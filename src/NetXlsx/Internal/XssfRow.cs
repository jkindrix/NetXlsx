using System;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfRow : IRow
{
    private readonly XssfWorkbook _workbook;
    private readonly XssfSheet _sheet;
    private readonly XSSFRow _underlying;
    private readonly int _row1;

    public XssfRow(XssfWorkbook workbook, XssfSheet sheet, XSSFRow underlying)
    {
        _workbook = workbook;
        _sheet = sheet;
        _underlying = underlying;
        _row1 = underlying.RowNum + 1;
    }

    public int Index
    {
        get { _workbook.ThrowIfDisposed(); return _row1; }
    }

    public ISheet Sheet
    {
        get { _workbook.ThrowIfDisposed(); return _sheet; }
    }

    public ICell Cell(int column)
    {
        _workbook.ThrowIfDisposed();
        int colCap = Math.Min(_workbook.Options.MaxColsPerSheet, CellAddress.MaxColumn);
        if (column < 1 || column > colCap)
            throw new ArgumentOutOfRangeException(nameof(column), column,
                $"column must be in [1, {colCap}]");

        int col0 = column - 1;
        var npoiCell = (XSSFCell?)_underlying.GetCell(col0) ?? (XSSFCell)_underlying.CreateCell(col0);
        return new XssfCell(_workbook, npoiCell, _row1, column);
    }

    public ICell this[int column] => Cell(column);

    public ICell this[string columnLetter]
    {
        get
        {
            _workbook.ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(columnLetter);
            // CellAddress.Format takes (row, col); we reverse via parse of
            // "<letter>1" to get the column index (cheap, reuses the parser).
            var (_, col) = CellAddress.Parse($"{columnLetter}1");
            return Cell(col);
        }
    }

    public IRow Set(int column, string value)   { Cell(column).SetString(value);   return this; }
    public IRow Set(int column, double value)   { Cell(column).SetNumber(value);   return this; }
    public IRow Set(int column, decimal value)  { Cell(column).SetNumber(value);   return this; }
    public IRow Set(int column, int value)      { Cell(column).SetNumber(value);   return this; }
    public IRow Set(int column, long value)     { Cell(column).SetNumber(value);   return this; }
    public IRow Set(int column, bool value)     { Cell(column).SetBool(value);     return this; }
    public IRow Set(int column, DateTime value) { Cell(column).SetDate(value);     return this; }
    public IRow Set(int column, DateOnly value) { Cell(column).SetDate(value);     return this; }
    public IRow Set(int column, TimeOnly value) { Cell(column).SetTime(value);     return this; }
    public IRow Set(int column, TimeSpan value) { Cell(column).SetDuration(value); return this; }

    public float HeightInPoints
    {
        get { _workbook.ThrowIfDisposed(); return _underlying.HeightInPoints; }
        set { _workbook.ThrowIfDisposed(); _underlying.HeightInPoints = value; }
    }

    public bool Hidden
    {
        get { _workbook.ThrowIfDisposed(); return _underlying.ZeroHeight; }
        set { _workbook.ThrowIfDisposed(); _underlying.ZeroHeight = value; }
    }

    public XSSFRow Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _underlying; }
    }
}
