// I-82 engine swap — Open XML SDK-backed IColumn (styles slice + I-84 AutoSize).
//
// Width / Hidden / SetDefaultStyle write the worksheet's <col> entry. The width
// unit is OOXML's "character width" (same unit IColumn.WidthUnits exposes) and is
// stored verbatim in @width; round-trips exactly within this engine. AutoSize
// measures with the embedded font-metric tables (OoxmlFontMetrics, decision
// I-84) and reproduces NPOI SheetUtil's width formula — see OoxmlFontMetrics.cs
// for the math and its documented approximations.

using System;
using System.Collections.Generic;

namespace NetXlsx;

internal sealed class OoxmlColumn : IColumn
{
    private readonly OoxmlSheet _sheet;
    private readonly int _index;

    internal OoxmlColumn(OoxmlSheet sheet, int index)
    {
        _sheet = sheet;
        _index = index;
    }

    public int Index { get { _sheet.WorkbookInternal.ThrowIfDisposed(); return _index; } }
    public string Letter { get { _sheet.WorkbookInternal.ThrowIfDisposed(); return CellAddress.FormatColumn(_index); } }
    public ISheet Sheet { get { _sheet.WorkbookInternal.ThrowIfDisposed(); return _sheet; } }

    // Excel's default column width in character units when none is specified.
    private const double DefaultWidthUnits = 8.43;

    private OoxmlWorkbook Wb => _sheet.WorkbookInternal;

    public bool Hidden
    {
        get { Wb.ThrowIfDisposed(); return _sheet.FindColumn(_index)?.Hidden?.Value ?? false; }
        set
        {
            Wb.ThrowIfDisposed();
            var col = _sheet.GetOrCreateColumn(_index);
            col.Hidden = value ? true : null;
            EnsureWidth(col);
        }
    }

    public double WidthUnits
    {
        get { Wb.ThrowIfDisposed(); return _sheet.FindColumn(_index)?.Width?.Value ?? DefaultWidthUnits; }
        set
        {
            Wb.ThrowIfDisposed();
            if (double.IsNaN(value) || value < 0)
                throw new ArgumentOutOfRangeException(nameof(value), value, "width must be non-negative and not NaN");
            var col = _sheet.GetOrCreateColumn(_index);
            col.Width = value;
            col.CustomWidth = true;
        }
    }

    public IColumn Width(double units)
    {
        WidthUnits = units;
        return this;
    }

    public IColumn SetDefaultStyle(CellStyle style)
    {
        Wb.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(style);
        var col = _sheet.GetOrCreateColumn(_index);
        col.Style = Wb.StylePool.GetOrCreate(style);
        EnsureWidth(col);
        return this;
    }

    public IColumn ForEachPopulated(Action<ICell> apply)
    {
        Wb.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(apply);
        foreach (var cell in _sheet.EnumeratePopulated(1, _index, CellAddress.MaxRow, _index))
            apply(cell);
        return this;
    }

    // A <col> entry is invalid without a width; when only hidden/style is set,
    // Excel still expects a width attribute, so default it if absent.
    private static void EnsureWidth(DocumentFormat.OpenXml.Spreadsheet.Column col)
    {
        if (col.Width is null) col.Width = DefaultWidthUnits;
    }

    // ---- AutoSize (I-84 — embedded font metrics) ---------------------------

    // Reproduces NPOI 2.7.3 SheetUtil.GetColumnWidth + XSSFSheet.AutoSizeColumn
    // (oracle-dumped 2026-06-03), measured with the embedded metric tables:
    //   - AutoSizeColumn(col) runs with useMergedCells=false → cells inside a
    //     merged region are skipped outright;
    //   - each \n-separated line measures independently;
    //   - cellWidth = (round(linePx) + 5) / defaultCharWidth * 1.05 + indent;
    //   - defaultCharWidth = ceil(ink('0') of font 0, px @96dpi);
    //   - nothing measurable → no-op (no <col> written), like NPOI;
    //   - result capped at 255 units; <col> gets bestFit + customWidth.
    // A font outside the embedded set throws MissingFontException (fail loud,
    // never silently-wrong widths — I-3's contract, now environment-free).
    public IColumn AutoSize()
    {
        Wb.ThrowIfDisposed();

        List<(int R1, int R2)>? merged = null;
        foreach (var range in _sheet.MergedRanges)
        {
            var (r1, c1, r2, c2) = CellAddress.ParseRange(range);
            if (c1 <= _index && _index <= c2)
                (merged ??= new()).Add((r1, r2));
        }

        double width = -1;
        int defaultCharWidth = 0;   // resolved on the first measured cell

        foreach (var cell in _sheet.EnumeratePopulated(1, _index, CellAddress.MaxRow, _index))
        {
            if (merged is not null && InMergedRows(merged, cell.RowIndex)) continue;
            string? text = cell.TextForMeasurement();
            if (text is null) continue;

            var style = Wb.StylePool.AutoSizeStyleOf(cell.XfIndex);
            var table = OoxmlFontMetrics.Resolve(style.FontName, style.Bold, style.Italic);
            if (defaultCharWidth == 0)
            {
                var def = Wb.StylePool.AutoSizeStyleOf(0);
                var defTable = OoxmlFontMetrics.Resolve(def.FontName, def.Bold, def.Italic);
                defaultCharWidth = OoxmlFontMetrics.DefaultCharWidthPx(defTable, def.FontSize);
            }

            foreach (var line in text.Split('\n'))
            {
                double raw = OoxmlFontMetrics.MeasureLineRawPx(table, line, style.FontSize);
                double actual;
                if (style.RotationRaw != 0)
                {
                    // NPOI's rotation trig, over the approximated line height
                    // (raw OOXML 0–180 rotation value, as NPOI uses it).
                    double angle = style.RotationRaw * 2.0 * Math.PI / 360.0;
                    double h = OoxmlFontMetrics.ApproxLineHeightRawPx(table, style.FontSize);
                    actual = Math.Round(
                        Math.Abs(h * Math.Sin(angle)) + Math.Abs(raw * Math.Cos(angle)),
                        0, MidpointRounding.ToEven);
                }
                else
                {
                    actual = Math.Round(raw, 0, MidpointRounding.ToEven);
                }

                double w = (actual + 5) / defaultCharWidth * 1.05 + style.Indent;
                if (w > width) width = w;
            }
        }

        if (width < 0) return this;     // nothing measurable — no-op (NPOI parity)

        if (width > 255) width = 255;   // Excel's per-cell width ceiling
        var col = _sheet.GetOrCreateColumn(_index);
        col.Width = width;
        col.CustomWidth = true;
        col.BestFit = true;
        return this;

        static bool InMergedRows(List<(int R1, int R2)> regions, int row)
        {
            foreach (var (r1, r2) in regions)
                if (row >= r1 && row <= r2) return true;
            return false;
        }
    }
}
