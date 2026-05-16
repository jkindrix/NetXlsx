using System;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfSheet : ISheet
{
    private readonly XssfWorkbook _workbook;
    private readonly XSSFSheet _underlying;

    public XssfSheet(XssfWorkbook workbook, XSSFSheet underlying)
    {
        _workbook = workbook;
        _underlying = underlying;
    }

    public string Name
    {
        get { _workbook.ThrowIfDisposed(); return _underlying.SheetName; }
    }

    public IWorkbook Workbook
    {
        get { _workbook.ThrowIfDisposed(); return _workbook; }
    }

    public ICell this[string a1]
    {
        get
        {
            _workbook.ThrowIfDisposed();
            var (row1, col1) = CellAddress.Parse(a1);
            return GetOrCreateCell(row1, col1);
        }
    }

    public ICell this[int row, int column]
    {
        get
        {
            _workbook.ThrowIfDisposed();
            if (row < 1 || row > CellAddress.MaxRow)
                throw new ArgumentOutOfRangeException(nameof(row), row,
                    $"row must be in [1, {CellAddress.MaxRow}]");
            if (column < 1 || column > CellAddress.MaxColumn)
                throw new ArgumentOutOfRangeException(nameof(column), column,
                    $"column must be in [1, {CellAddress.MaxColumn}]");
            return GetOrCreateCell(row, column);
        }
    }

    public IRow AppendRow()
    {
        _workbook.ThrowIfDisposed();
        int next0 = _underlying.PhysicalNumberOfRows == 0 ? 0 : _underlying.LastRowNum + 1;
        if (next0 + 1 > CellAddress.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(AppendRow),
                $"appending would exceed Excel's row limit of {CellAddress.MaxRow}");
        var npoiRow = (XSSFRow)_underlying.CreateRow(next0);
        return new XssfRow(_workbook, this, npoiRow);
    }

    public IRow Row(int index)
    {
        _workbook.ThrowIfDisposed();
        if (index < 1 || index > CellAddress.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(index), index,
                $"row index must be in [1, {CellAddress.MaxRow}]");
        int row0 = index - 1;
        var npoiRow = (XSSFRow?)_underlying.GetRow(row0) ?? (XSSFRow)_underlying.CreateRow(row0);
        return new XssfRow(_workbook, this, npoiRow);
    }

    public XSSFSheet Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _underlying; }
    }

    private XssfCell GetOrCreateCell(int row1, int col1)
    {
        // NPOI rows/cols are 0-based; ours are 1-based per decision #3.
        int row0 = row1 - 1;
        int col0 = col1 - 1;

        // Materialize on access (decision #40 — every cell exists).
        var npoiRow = (XSSFRow?)_underlying.GetRow(row0) ?? (XSSFRow)_underlying.CreateRow(row0);
        var npoiCell = (XSSFCell?)npoiRow.GetCell(col0) ?? (XSSFCell)npoiRow.CreateCell(col0);
        return new XssfCell(_workbook, npoiCell, row1, col1);
    }
}
