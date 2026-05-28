using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfConnector : IConnector
{
    private readonly XssfSheet _sheet;
    private readonly XSSFConnector _underlying;
    private readonly ConnectorType _type;

    internal XssfConnector(XssfSheet sheet, XSSFConnector underlying, ConnectorType type)
    {
        _sheet = sheet;
        _underlying = underlying;
        _type = type;
    }

    public ISheet Sheet => _sheet;
    public ConnectorType Type => _type;
    public XSSFConnector Underlying => _underlying;
}
