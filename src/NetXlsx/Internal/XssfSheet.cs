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
            // NPOI rows/cols are 0-based; ours are 1-based per decision #3.
            int row0 = row1 - 1;
            int col0 = col1 - 1;

            // Materialize on access (decision #40 — every cell exists).
            var npoiRow = (XSSFRow?)_underlying.GetRow(row0) ?? (XSSFRow)_underlying.CreateRow(row0);
            var npoiCell = (XSSFCell?)npoiRow.GetCell(col0) ?? (XSSFCell)npoiRow.CreateCell(col0);
            return new XssfCell(_workbook, npoiCell, row1, col1);
        }
    }

    public XSSFSheet Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _underlying; }
    }
}
