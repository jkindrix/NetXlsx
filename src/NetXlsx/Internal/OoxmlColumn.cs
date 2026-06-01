// I-82 engine swap — Open XML SDK-backed IColumn (styles slice).
//
// Width / Hidden / SetDefaultStyle write the worksheet's <col> entry. The width
// unit is OOXML's "character width" (same unit IColumn.WidthUnits exposes) and is
// stored verbatim in @width; round-trips exactly within this engine. AutoSize
// stays deferred — it needs font-metric measurement, which the SDK engine does
// not carry (it lands with the font-metrics work, not this slice).

using System;
using System.Runtime.CompilerServices;

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

    // ---- Deferred (font-metrics work; see I-82) ----------------------------

    private static NotImplementedException NotYet([CallerMemberName] string? member = null)
        => new(
            $"IColumn.{member} is not yet implemented on the Open XML SDK engine " +
            "(I-82 engine swap). AutoSize needs font-metric measurement, which lands " +
            "with the font-metrics work; track the swap in docs/design.md (I-82).");

    public IColumn AutoSize() => throw NotYet();
}
