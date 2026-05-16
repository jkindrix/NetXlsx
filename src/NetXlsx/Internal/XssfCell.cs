using System;
using System.Globalization;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfCell : ICell
{
    private readonly XssfWorkbook _workbook;
    private readonly XSSFCell _underlying;
    private readonly int _row1;
    private readonly int _col1;

    public XssfCell(XssfWorkbook workbook, XSSFCell underlying, int row1Based, int col1Based)
    {
        _workbook = workbook;
        _underlying = underlying;
        _row1 = row1Based;
        _col1 = col1Based;
    }

    public string Address
    {
        get { _workbook.ThrowIfDisposed(); return CellAddress.Format(_row1, _col1); }
    }

    public int RowIndex
    {
        get { _workbook.ThrowIfDisposed(); return _row1; }
    }

    public int ColumnIndex
    {
        get { _workbook.ThrowIfDisposed(); return _col1; }
    }

    public CellKind Kind
    {
        get
        {
            _workbook.ThrowIfDisposed();
            return ClassifyKind(_underlying);
        }
    }

    private static CellKind ClassifyKind(XSSFCell cell)
    {
        switch (cell.CellType)
        {
            case CellType.Blank: return CellKind.Empty;
            case CellType.String: return CellKind.String;
            case CellType.Boolean: return CellKind.Bool;
            case CellType.Error: return CellKind.Error;
            case CellType.Formula: return CellKind.Formula;
            case CellType.Numeric:
                return DateUtil.IsCellDateFormatted(cell) ? CellKind.Date : CellKind.Number;
            default: return CellKind.Empty;
        }
    }

    public void SetString(string value)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(value);
        _underlying.SetCellValue(value);
    }

    public void SetNumber(double value)
    {
        _workbook.ThrowIfDisposed();
        _underlying.SetCellValue(value);
    }

    public void SetNumber(decimal value)
    {
        _workbook.ThrowIfDisposed();
        // Decision I3.6 / §7.4: stored as IEEE-754 double; precision loss
        // possible for > ~15 significant digits.
        _underlying.SetCellValue((double)value);
    }

    public void SetNumber(int value)
    {
        _workbook.ThrowIfDisposed();
        _underlying.SetCellValue((double)value);
    }

    public void SetNumber(long value)
    {
        _workbook.ThrowIfDisposed();
        // Values > 2^53 lose precision when stored as IEEE-754 double.
        _underlying.SetCellValue((double)value);
    }

    public void SetBool(bool value)
    {
        _workbook.ThrowIfDisposed();
        _underlying.SetCellValue(value);
    }

    public void Clear()
    {
        _workbook.ThrowIfDisposed();
        _underlying.SetBlank();
    }

    public string GetString()
    {
        _workbook.ThrowIfDisposed();
        // Per design §7.10:
        //   Empty -> ""; String -> stored verbatim; Number -> invariant string;
        //   Bool -> "TRUE" / "FALSE" (invariant); Formula -> cached result;
        //   Error -> error code text.
        switch (_underlying.CellType)
        {
            case CellType.Blank: return string.Empty;
            case CellType.String: return _underlying.StringCellValue ?? string.Empty;
            case CellType.Boolean: return _underlying.BooleanCellValue ? "TRUE" : "FALSE";
            case CellType.Numeric:
                return _underlying.NumericCellValue.ToString("G17", CultureInfo.InvariantCulture);
            case CellType.Error:
                return FormulaError.ForInt(_underlying.ErrorCellValue).String;
            case CellType.Formula:
                return GetFormulaCachedAsString();
            default: return string.Empty;
        }
    }

    private string GetFormulaCachedAsString()
    {
        switch (_underlying.CachedFormulaResultType)
        {
            case CellType.String: return _underlying.StringCellValue ?? string.Empty;
            case CellType.Boolean: return _underlying.BooleanCellValue ? "TRUE" : "FALSE";
            case CellType.Numeric:
                return _underlying.NumericCellValue.ToString("G17", CultureInfo.InvariantCulture);
            case CellType.Error:
                return FormulaError.ForInt(_underlying.ErrorCellValue).String;
            default: return string.Empty;
        }
    }

    public double? GetNumber()
    {
        _workbook.ThrowIfDisposed();
        switch (_underlying.CellType)
        {
            case CellType.Numeric: return _underlying.NumericCellValue;
            case CellType.Boolean: return _underlying.BooleanCellValue ? 1.0 : 0.0;
            case CellType.Formula:
                return _underlying.CachedFormulaResultType == CellType.Numeric
                    ? _underlying.NumericCellValue
                    : (double?)null;
            default: return null;
        }
    }

    public bool? GetBool()
    {
        _workbook.ThrowIfDisposed();
        return _underlying.CellType switch
        {
            CellType.Boolean => _underlying.BooleanCellValue,
            CellType.Formula when _underlying.CachedFormulaResultType == CellType.Boolean
                => _underlying.BooleanCellValue,
            _ => null,
        };
    }

    public XSSFCell Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _underlying; }
    }
}
