using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfShape : IShape
{
    private readonly XssfSheet _sheet;
    private readonly XSSFSimpleShape _underlying;
    private readonly ShapeType _type;

    internal XssfShape(XssfSheet sheet, XSSFSimpleShape underlying, ShapeType type)
    {
        _sheet = sheet;
        _underlying = underlying;
        _type = type;
    }

    public ISheet Sheet => _sheet;
    public ShapeType Type => _type;
    public XSSFSimpleShape Underlying => _underlying;
}
