// XssfCell — core. Fields, ctor, identity getters, kind classification,
// clear, .Underlying escape hatch, and the default-style helper.
// Value setters/getters split to XssfCell.Values.cs; styling to
// XssfCell.Style.cs; comments + hyperlinks to XssfCell.Annotations.cs.

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed partial class XssfCell : ICell
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

    public void Clear()
    {
        _workbook.ThrowIfDisposed();
        _underlying.SetBlank();
    }

    public XSSFCell Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _underlying; }
    }

    /// <summary>
    /// Applies <paramref name="defaultStyle"/> to the cell when it currently
    /// carries no explicit style. Per decision I-18: a user-set style is
    /// preserved. The workbook-default style has index 0; any explicit
    /// style has a higher index.
    /// </summary>
    private void ApplyDefaultStyleIfUnstyled(ICellStyle defaultStyle)
    {
        var current = _underlying.CellStyle;
        if (current is null || current.Index == 0)
        {
            _underlying.CellStyle = defaultStyle;
        }
    }
}
