// I-82 engine swap — Open XML SDK-backed IShape (drawings slice, shapes/connectors).
//
// A shape on the SDK engine. IShape's surface is intentionally minimal (Sheet, Type,
// Underlying) — there is no ISheet.Shapes read-back, so this is a lightweight wrapper
// over the value passed to OoxmlSheet.AddShape; the geometry lives in the xdr:sp the
// add path built, which the escape hatch (Underlying, #32 / I-82) hands out live.

using System;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlShape : IShape
{
    private readonly OoxmlWorkbook _workbook;
    private readonly OoxmlSheet _sheet;
    private readonly ShapeType _type;
    private readonly XDR.Shape _shape;

    internal OoxmlShape(OoxmlWorkbook workbook, OoxmlSheet sheet, ShapeType type, XDR.Shape shape)
    {
        _workbook = workbook;
        _sheet = sheet;
        _type = type;
        _shape = shape;
    }

    public ISheet Sheet { get { _workbook.ThrowIfDisposed(); return _sheet; } }
    public ShapeType Type { get { _workbook.ThrowIfDisposed(); return _type; } }

    // Escape hatch (#32 / I-82): the live xdr:sp element. Disposal first.
    public XDR.Shape Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _shape; }
    }
}
