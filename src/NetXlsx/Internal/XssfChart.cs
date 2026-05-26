using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfChart : IChart
{
    private readonly XssfSheet _sheet;
    private readonly XSSFChart _underlying;
    private readonly ChartType _type;

    internal XssfChart(XssfSheet sheet, XSSFChart underlying, ChartType type)
    {
        _sheet = sheet;
        _underlying = underlying;
        _type = type;
    }

    public ISheet Sheet => _sheet;
    public ChartType Type => _type;

    public void SetTitle(string title) => _underlying.SetTitle(title);

    public XSSFChart Underlying => _underlying;
}
