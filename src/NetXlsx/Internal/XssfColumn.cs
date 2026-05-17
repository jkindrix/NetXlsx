using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfColumn : IColumn
{
    private readonly XssfWorkbook _workbook;
    private readonly XssfSheet _sheet;
    private readonly int _col1;     // 1-based
    private readonly int _col0;     // 0-based for NPOI

    public XssfColumn(XssfWorkbook workbook, XssfSheet sheet, int column1)
    {
        _workbook = workbook;
        _sheet = sheet;
        _col1 = column1;
        _col0 = column1 - 1;
    }

    public int Index { get { _workbook.ThrowIfDisposed(); return _col1; } }

    public string Letter { get { _workbook.ThrowIfDisposed(); return CellAddress.FormatColumn(_col1); } }

    public ISheet Sheet { get { _workbook.ThrowIfDisposed(); return _sheet; } }

    public bool Hidden
    {
        get { _workbook.ThrowIfDisposed(); return _sheet.Underlying.IsColumnHidden(_col0); }
        set { _workbook.ThrowIfDisposed(); _sheet.Underlying.SetColumnHidden(_col0, value); }
    }

    public double WidthUnits
    {
        get { _workbook.ThrowIfDisposed(); return _sheet.Underlying.GetColumnWidth(_col0) / 256.0; }
        set
        {
            _workbook.ThrowIfDisposed();
            ValidateWidth(value);
            _sheet.Underlying.SetColumnWidth(_col0, (int)Math.Round(value * 256.0));
        }
    }

    public IColumn Width(double units)
    {
        WidthUnits = units;
        return this;
    }

    public IColumn AutoSize()
    {
        _workbook.ThrowIfDisposed();
        try
        {
            _sheet.Underlying.AutoSizeColumn(_col0);
        }
        catch (Exception ex) when (IsFontFailure(ex))
        {
            throw new MissingFontException(ex);
        }
        return this;
    }

    public IColumn ForEachPopulated(Action<ICell> apply)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(apply);

        var npoiSheet = _sheet.Underlying;
        int last0 = npoiSheet.LastRowNum;
        for (int r0 = 0; r0 <= last0; r0++)
        {
            var npoiRow = (XSSFRow?)npoiSheet.GetRow(r0);
            if (npoiRow is null) continue;
            var npoiCell = (XSSFCell?)npoiRow.GetCell(_col0);
            if (npoiCell is null) continue;
            var cell = new XssfCell(_workbook, npoiCell, r0 + 1, _col1);
            if (cell.Kind == CellKind.Empty) continue;
            apply(cell);
        }
        return this;
    }

    public IColumn SetDefaultStyle(CellStyle style)
    {
        _workbook.ThrowIfDisposed();
        var npoiStyle = _workbook.StylePool.GetOrCreate(style);
        _sheet.Underlying.SetDefaultColumnStyle(_col0, npoiStyle);
        return this;
    }

    private static void ValidateWidth(double units)
    {
        if (double.IsNaN(units) || units < 0)
            throw new ArgumentOutOfRangeException(nameof(units), units, "width must be non-negative and not NaN");
    }

    // NPOI's AutoSizeColumn ultimately calls into SixLabors.Fonts /
    // System.Drawing for font metrics. When those can't resolve a font,
    // failures surface as a small set of exception types — we translate
    // them to MissingFontException with installation guidance (I3).
    private static bool IsFontFailure(Exception ex)
    {
        for (Exception? e = ex; e is not null; e = e.InnerException)
        {
            var name = e.GetType().FullName ?? string.Empty;
            if (name.StartsWith("SixLabors.Fonts.", StringComparison.Ordinal)) return true;
            if (name == "System.Drawing.SystemFontsException") return true;
            if (e is TypeInitializationException) return true;
            if (e is System.IO.FileNotFoundException fnf
                && (fnf.FileName?.Contains("libgdiplus", StringComparison.OrdinalIgnoreCase) == true
                    || fnf.Message.Contains("font", StringComparison.OrdinalIgnoreCase)))
                return true;
            if (e is System.IO.IOException && e.Message.Contains("font", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }
}
