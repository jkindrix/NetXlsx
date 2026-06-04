using NPOI.SS.UserModel;

namespace NetXlsx;

internal sealed class XssfNamedRange : INamedRange
{
    private readonly XssfWorkbook _workbook;
    private readonly IName _underlying;

    public XssfNamedRange(XssfWorkbook workbook, IName underlying)
    {
        _workbook = workbook;
        _underlying = underlying;
    }

    public string Name { get { _workbook.ThrowIfDisposed(); return _underlying.NameName; } }

    public string Formula { get { _workbook.ThrowIfDisposed(); return _underlying.RefersToFormula; } }

    public string? SheetScope
    {
        get
        {
            _workbook.ThrowIfDisposed();
            int idx = _underlying.SheetIndex;
            return idx < 0 ? null : _workbook.Npoi.GetSheetName(idx);
        }
    }
}
