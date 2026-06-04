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

    // v2.0.0 (I-82): SDK-typed hatch; nothing to expose on the NPOI engine.
    public DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape Underlying => throw new System.NotSupportedException(
        "IShape.Underlying (xdr:sp) is not available on the retired NPOI engine.");
}
