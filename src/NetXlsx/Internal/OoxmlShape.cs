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

    // ---- Removed-handle access guard (I-91 slice 2) -----------------------
    // The drawing-layer twin of the OoxmlTable retrofit (S14): after
    // OoxmlSheet.RemoveShape detaches this shape's anchor, every public member
    // throws InvalidOperationException — distinct from the disposed-workbook
    // ObjectDisposedException. The flag is one-way.
    private bool _removed;

    internal void MarkRemoved() => _removed = true;

    // The live xdr:sp element, for RemoveShape's anchor match (no liveness
    // guard — internal engine use only).
    internal XDR.Shape ShapeElement => _shape;

    // Disposal first so a disposed workbook still surfaces
    // ObjectDisposedException; a live workbook with this shape removed surfaces
    // InvalidOperationException.
    internal void ThrowIfUnusable()
    {
        _workbook.ThrowIfDisposed();
        if (_removed)
            throw new InvalidOperationException(
                "this shape has been removed from its sheet.");
    }

    public ISheet Sheet { get { ThrowIfUnusable(); return _sheet; } }
    public ShapeType Type { get { ThrowIfUnusable(); return _type; } }

    // Escape hatch (#32 / I-82): the live xdr:sp element. Disposal first.
    public XDR.Shape Underlying
    {
        get { ThrowIfUnusable(); return _shape; }
    }
}
