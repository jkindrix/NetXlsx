// I-82 engine swap — Open XML SDK-backed IShape (drawings slice, shapes/connectors).
//
// A shape on the SDK engine. IShape's surface is intentionally minimal (Sheet, Type,
// Underlying) — there is no ISheet.Shapes read-back, so this is a lightweight wrapper
// over the value passed to OoxmlSheet.AddShape; the geometry lives in the xdr:sp the
// add path built. The NPOI escape hatch (Underlying -> XSSFSimpleShape) throws
// NotSupportedException, the same divergence as OoxmlPicture/OoxmlConnector: there is
// no NPOI object behind the SDK engine. The SDK package is reachable via
// IWorkbook.OpenXmlDocument.

using System;

namespace NetXlsx;

internal sealed class OoxmlShape : IShape
{
    private readonly OoxmlWorkbook _workbook;
    private readonly OoxmlSheet _sheet;
    private readonly ShapeType _type;

    internal OoxmlShape(OoxmlWorkbook workbook, OoxmlSheet sheet, ShapeType type)
    {
        _workbook = workbook;
        _sheet = sheet;
        _type = type;
    }

    public ISheet Sheet { get { _workbook.ThrowIfDisposed(); return _sheet; } }
    public ShapeType Type { get { _workbook.ThrowIfDisposed(); return _type; } }

    // Escape-hatch divergence (I-82): no NPOI shape exists on the SDK engine.
    public NPOI.XSSF.UserModel.XSSFSimpleShape Underlying => throw new NotSupportedException(
        "IShape.Underlying (NPOI XSSFSimpleShape) is not available on the Open XML SDK " +
        "engine (I-82). Use IWorkbook.OpenXmlDocument for the SDK escape hatch.");
}
