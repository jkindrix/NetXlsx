// I-82 engine swap — Open XML SDK-backed shapes + connectors (drawings slice).
//
// Shapes and connectors are anchor children of the same xdr:wsDr root the pictures
// sub-slice already builds (GetOrCreateDrawing returns it; the worksheet <drawing>
// child is routed through OoxmlSchemaOrder). xdr:wsDr is NOT a strict-ordered
// container — anchors append freely — so shapes/connectors do not hit SDK-quirk #8
// inside it (only the worksheet <drawing> child does, which the pictures path owns).
//
// The geometry mirrors the NPOI engine exactly (the parity oracle — see lesson #6:
// schema-valid != positioned-correctly, so the SDK output is asserted against the
// NPOI engine's ST_ShapeType preset, EMU markers, and <a:ln> props, not just
// OpenXmlValidator). Captured from the NPOI engine's emitted xl/drawings/drawingN.xml:
//
//   Shape (xdr:sp) — twoCellAnchor, end cell EXCLUSIVE (from c1-1/r1-1, to c2/r2,
//     same convention as a picture's two-cell anchor). spPr: xfrm(off=0,ext=0) +
//     prstGeom(rect|roundRect|ellipse|line|triangle|diamond) + (solidFill|noFill) +
//     optional <a:ln><a:solidFill>. A minimal txBody follows spPr (CT_Shape carries
//     text). No <xdr:style> block. IShape exposes only Sheet/Type, so this is a
//     write-only fidelity surface (there is no ISheet.Shapes read-back).
//
//   Connector (xdr:cxnSp) — twoCellAnchor, end cell INCLUSIVE (from c1-1/r1-1, to
//     c2-1/r2-1, preserving per-end EMU offsets dx1/dy1/dx2/dy2). This differs from
//     shapes/pictures: a connector's NPOI anchor maps the end cell with -1, so a
//     connector that starts and ends in the same cell (I41->I41) round-trips ToCell
//     == FromCell. spPr: xfrm(flipH/flipV) + prstGeom(straightConnector1=96 /
//     bentConnector3=98 / curvedConnector3=102 — the public ConnectorType ordinals,
//     lesson #6) + optional <a:ln w=...> with solidFill + head/tail line ends. A
//     <xdr:style> block (lnRef idx=1/accent1, fillRef/effectRef idx=0, fontRef minor)
//     always follows spPr, matching NPOI — so LineStyleRefIndex reads 1 and
//     LineSchemeColor falls back to "accent1" exactly as on the NPOI engine.

using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    // EMU per point — <a:ln @w> is EMU; the public surface is points (lesson #6).
    private const long EmuPerPoint = 12700;

    public IShape AddShape(ShapeType type, string startCell, string endCell,
        Color? fillColor = null, Color? lineColor = null)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(startCell);
        ArgumentNullException.ThrowIfNull(endCell);

        var (r1, c1) = CellAddress.Parse(startCell);
        var (r2, c2) = CellAddress.Parse(endCell);

        var (_, root) = GetOrCreateDrawing();
        uint id = NextShapeId(root);

        var spPr = new XDR.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 0L, Cy = 0L }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = ToPreset(type) });
        // CT_ShapeProperties order: xfrm, geometry, fill, ln.
        spPr.Append(fillColor is { } fc
            ? new A.SolidFill(new A.RgbColorModelHex { Val = Hex(fc) })
            : (DocumentFormat.OpenXml.OpenXmlElement)new A.NoFill());
        if (lineColor is { } lc)
            spPr.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = Hex(lc) })));

        var shape = new XDR.Shape(
            new XDR.NonVisualShapeProperties(
                new XDR.NonVisualDrawingProperties { Id = id, Name = $"Shape {id}" },
                new XDR.NonVisualShapeDrawingProperties()),
            spPr,
            new XDR.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" })));

        // Shape anchor end cell is exclusive (NPOI: XSSFClientAnchor(...,c2,r2)).
        root.AppendChild(new XDR.TwoCellAnchor(
            Marker(c1 - 1, 0, r1 - 1, 0),
            ToMarker(c2, 0, r2, 0),
            shape,
            new XDR.ClientData()));

        return new OoxmlShape(_workbook, this, type);
    }

    public IConnector AddConnector(ConnectorType type, string startCell, string endCell,
        Color? lineColor = null,
        int dx1 = 0, int dy1 = 0, int dx2 = 0, int dy2 = 0,
        bool flipH = false, bool flipV = false,
        ConnectorEnd headEnd = ConnectorEnd.None, ConnectorEnd tailEnd = ConnectorEnd.None,
        double? lineWidthPoints = null)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(startCell);
        ArgumentNullException.ThrowIfNull(endCell);

        var (r1, c1) = CellAddress.Parse(startCell);
        var (r2, c2) = CellAddress.Parse(endCell);

        var (_, root) = GetOrCreateDrawing();
        uint id = NextShapeId(root);

        var xfrm = new A.Transform2D(
            new A.Offset { X = 0L, Y = 0L },
            new A.Extents { Cx = 0L, Cy = 0L });
        if (flipH) xfrm.HorizontalFlip = true;
        if (flipV) xfrm.VerticalFlip = true;

        var spPr = new XDR.ShapeProperties(
            xfrm,
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = ToPreset(type) });

        // <a:ln> only when the caller pins a line property (NPOI omits it otherwise).
        if (lineColor is not null || lineWidthPoints is not null
            || headEnd != ConnectorEnd.None || tailEnd != ConnectorEnd.None)
        {
            var ln = new A.Outline();
            if (lineWidthPoints is { } w)
                ln.Width = (int)Math.Round(w * EmuPerPoint);
            if (lineColor is { } lc)
                ln.Append(new A.SolidFill(new A.RgbColorModelHex { Val = Hex(lc) }));
            // CT_LineProperties order: fill, ..., headEnd, tailEnd.
            if (headEnd != ConnectorEnd.None)
                ln.Append(new A.HeadEnd { Type = ToLineEnd(headEnd) });
            if (tailEnd != ConnectorEnd.None)
                ln.Append(new A.TailEnd { Type = ToLineEnd(tailEnd) });
            spPr.Append(ln);
        }

        var cxnSp = new XDR.ConnectionShape(
            new XDR.NonVisualConnectionShapeProperties(
                new XDR.NonVisualDrawingProperties { Id = id, Name = $"Connector {id}" },
                new XDR.NonVisualConnectorShapeDrawingProperties()),
            spPr,
            DefaultConnectorStyle());

        // Connector anchor end cell is inclusive (NPOI: XSSFClientAnchor(...,c2-1,r2-1)).
        var anchor = new XDR.TwoCellAnchor(
            Marker(c1 - 1, dx1, r1 - 1, dy1),
            ToMarker(c2 - 1, dx2, r2 - 1, dy2),
            cxnSp,
            new XDR.ClientData());
        root.AppendChild(anchor);

        return OoxmlConnector.FromElement(_workbook, this, anchor, cxnSp);
    }

    public IReadOnlyList<IConnector> Connectors
    {
        get
        {
            _workbook.ThrowIfDisposed();
            var result = new List<IConnector>();
            var drawingEl = Worksheet.GetFirstChild<S.Drawing>();
            if (drawingEl?.Id?.Value is not string rid) return result;
            if (_worksheetPart.GetPartById(rid) is not DrawingsPart dp) return result;
            var root = dp.WorksheetDrawing;
            if (root is null) return result;

            foreach (var child in root.ChildElements)
            {
                if (child is XDR.TwoCellAnchor anchor
                    && anchor.GetFirstChild<XDR.ConnectionShape>() is { } cxn)
                {
                    result.Add(OoxmlConnector.FromElement(_workbook, this, anchor, cxn));
                }
            }
            return result;
        }
    }

    // ---- Connector style block (matches the NPOI engine for read-back parity) ----

    // NPOI emits this fixed style on every connector: lnRef idx=1 accent1, fillRef /
    // effectRef idx=0 accent1, fontRef minor tx1. LineStyleRefIndex reads lnRef/@idx
    // (=1) and LineSchemeColor falls back to lnRef/schemeClr (="accent1").
    private static XDR.ShapeStyle DefaultConnectorStyle() => new(
        new A.LineReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 1U },
        new A.FillReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 0U },
        new A.EffectReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 0U },
        new A.FontReference(new A.SchemeColor { Val = A.SchemeColorValues.Text1 }) { Index = A.FontCollectionIndexValues.Minor });

    // ---- Mappings -----------------------------------------------------------

    private static A.ShapeTypeValues ToPreset(ShapeType type) => type switch
    {
        ShapeType.Rectangle => A.ShapeTypeValues.Rectangle,
        ShapeType.RoundedRectangle => A.ShapeTypeValues.RoundRectangle,
        ShapeType.Ellipse => A.ShapeTypeValues.Ellipse,
        ShapeType.Line => A.ShapeTypeValues.Line,
        ShapeType.Triangle => A.ShapeTypeValues.Triangle,
        ShapeType.Diamond => A.ShapeTypeValues.Diamond,
        _ => A.ShapeTypeValues.Rectangle,
    };

    private static A.ShapeTypeValues ToPreset(ConnectorType type) => type switch
    {
        ConnectorType.Straight => A.ShapeTypeValues.StraightConnector1,
        ConnectorType.Bent => A.ShapeTypeValues.BentConnector3,
        ConnectorType.Curved => A.ShapeTypeValues.CurvedConnector3,
        _ => A.ShapeTypeValues.StraightConnector1,
    };

    private static A.LineEndValues ToLineEnd(ConnectorEnd end) => end switch
    {
        ConnectorEnd.Triangle => A.LineEndValues.Triangle,
        ConnectorEnd.Stealth => A.LineEndValues.Stealth,
        ConnectorEnd.Diamond => A.LineEndValues.Diamond,
        ConnectorEnd.Oval => A.LineEndValues.Oval,
        ConnectorEnd.Arrow => A.LineEndValues.Arrow,
        _ => A.LineEndValues.None,
    };

    private static string Hex(Color c) =>
        $"{c.R:X2}{c.G:X2}{c.B:X2}";

    // Marker(...) / ToMarker(...) / NextShapeId(...) are shared with the pictures
    // sub-slice (OoxmlSheet.Pictures.cs); GetOrCreateDrawing() owns the DrawingsPart
    // + the schema-ordered worksheet <drawing> child.

    // Parses a drawing-anchor marker leaf (xdr:col / xdr:row / xdr:colOff / xdr:rowOff),
    // failing loud on a missing or non-integer value rather than silently substituting 0
    // (decision I-83 — fail-loud parity with the NPOI engine, which rejects a corrupt
    // anchor on open). CT_Marker requires all four children, so an absent or unparseable
    // marker is genuine file corruption; a silent 0 would mis-place the drawing. Shared
    // with the pictures sub-slice (OoxmlSheet.Pictures.cs read-back) and connectors.
    internal static int ParseMarker(DocumentFormat.OpenXml.OpenXmlLeafTextElement? el) =>
        int.TryParse(el?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v)
            ? v
            : throw new MalformedFileException(
                $"drawing anchor marker '{el?.Text}' is not a valid integer");
}
