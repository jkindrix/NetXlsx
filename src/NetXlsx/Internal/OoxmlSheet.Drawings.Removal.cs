// I-82 engine swap — Open XML SDK-backed drawing-layer removal (I-91 removal
// family, slice 2 of 2; ledger R-11). The handle-based half: RemovePicture /
// RemoveChart / RemoveShape / RemoveConnector, mirroring OoxmlSheet.RemoveTable
// semantics EXACTLY — ArgumentException on a foreign handle (not an Ooxml* of
// ours) or a stale one (its anchor/part no longer belongs to this sheet's
// drawing), and the passed handle becomes a tombstone (MarkRemoved) whose
// members throw InvalidOperationException afterward.
//
// Cleanup contract (the design's drawing-layer bullet + [A-2026-06-11]):
//   1. Delete the handle's ANCHOR (oneCellAnchor/twoCellAnchor wrapping the
//      xdr:pic / xdr:sp / xdr:cxnSp / xdr:graphicFrame). Matched by the handle's
//      LIVE element (ReferenceEquals on the stored XDR element, or the chart's
//      ChartPart), never by coordinates — two pictures can share a cell.
//   2. SHARED-MEDIA REFCOUNT [A-2026-06-11]: a picture's image part is deleted
//      only when no OTHER surviving anchor's blip still resolves to it.
//      AddPicture does NOT dedup image parts today (EmbedImage calls
//      AddImagePart fresh every call), but the guard stays anyway — it is cheap,
//      it protects opened third-party files that DO share media parts, and no
//      surveyed library refcounts media. A chart's ChartPart is not shared, so
//      it goes with its anchor — but still through DeletePart, so its rel (and
//      any colors/style sub-parts, dropped as unreachable by the clone-based
//      Save) leave too. Shapes/connectors own no separate part.
//   3. Drop the DrawingsPart + the worksheet <drawing> rel when the LAST anchor
//      goes (the no-empty-artifact discipline — same family as S13's RemoveSheet
//      part teardown and S14's RemoveComment VML-part guard).

using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    public void RemovePicture(IPicture picture)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(picture);
        if (picture is not OoxmlPicture op)
            throw new ArgumentException(
                "Picture instance is not an OoxmlPicture — likely a foreign or mocked implementation.",
                nameof(picture));

        var (drawingsPart, root) = RequireDrawing(nameof(picture));
        var anchor = root.ChildElements.FirstOrDefault(
            a => IsAnchor(a) && a.Descendants<XDR.Picture>().Any(p => ReferenceEquals(p, op.PicElement)));
        if (anchor is null)
            throw StaleHandle(nameof(picture), "picture");

        // Resolve the image part this picture's blip embeds BEFORE detaching the
        // anchor, so the refcount sweep over the survivors is meaningful.
        ImagePart? imagePart = ResolveBlipImagePart(drawingsPart, op.PicElement);

        anchor.Remove();

        // [A-2026-06-11] shared-media refcount: delete the image part only when
        // no surviving anchor's blip still resolves to it (ReferenceEquals on
        // the part, so distinct rel ids pointing at one shared part also count).
        if (imagePart is not null && !AnyBlipReferences(root, drawingsPart, imagePart))
            drawingsPart.DeletePart(imagePart);

        TeardownDrawingIfEmpty(drawingsPart, root);
        op.MarkRemoved();
    }

    public void RemoveChart(IChart chart)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(chart);
        if (chart is not OoxmlChart oc)
            throw new ArgumentException(
                "Chart instance is not an OoxmlChart — likely a foreign or mocked implementation.",
                nameof(chart));

        var (drawingsPart, root) = RequireDrawing(nameof(chart));
        var anchor = root.ChildElements.FirstOrDefault(
            a => IsAnchor(a) && AnchorReferencesChart(drawingsPart, a, oc.PartRef));
        if (anchor is null)
            throw StaleHandle(nameof(chart), "chart");

        anchor.Remove();

        // Charts aren't shared — the ChartPart (and its colors/style sub-parts,
        // dropped as unreachable by Save) goes with its anchor. DeletePart so the
        // relationship leaves too.
        drawingsPart.DeletePart(oc.PartRef);

        TeardownDrawingIfEmpty(drawingsPart, root);
        oc.MarkRemoved();
    }

    public void RemoveShape(IShape shape)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(shape);
        if (shape is not OoxmlShape os)
            throw new ArgumentException(
                "Shape instance is not an OoxmlShape — likely a foreign or mocked implementation.",
                nameof(shape));

        var (drawingsPart, root) = RequireDrawing(nameof(shape));
        var anchor = root.ChildElements.FirstOrDefault(
            a => IsAnchor(a) && a.Descendants<XDR.Shape>().Any(s => ReferenceEquals(s, os.ShapeElement)));
        if (anchor is null)
            throw StaleHandle(nameof(shape), "shape");

        anchor.Remove();   // a shape owns no separate part
        TeardownDrawingIfEmpty(drawingsPart, root);
        os.MarkRemoved();
    }

    public void RemoveConnector(IConnector connector)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(connector);
        if (connector is not OoxmlConnector oc)
            throw new ArgumentException(
                "Connector instance is not an OoxmlConnector — likely a foreign or mocked implementation.",
                nameof(connector));

        var (drawingsPart, root) = RequireDrawing(nameof(connector));
        var anchor = root.ChildElements.FirstOrDefault(
            a => IsAnchor(a) && a.Descendants<XDR.ConnectionShape>().Any(c => ReferenceEquals(c, oc.ConnectionElement)));
        if (anchor is null)
            throw StaleHandle(nameof(connector), "connector");

        anchor.Remove();   // a connector owns no separate part
        TeardownDrawingIfEmpty(drawingsPart, root);
        oc.MarkRemoved();
    }

    // ---- shared drawing-removal helpers -------------------------------------

    // WorksheetDrawing's children are anchors (oneCellAnchor/twoCellAnchor/
    // absoluteAnchor); only those count toward "is the drawing now empty".
    private static bool IsAnchor(OpenXmlElement e)
        => e is XDR.OneCellAnchor or XDR.TwoCellAnchor or XDR.AbsoluteAnchor;

    // The sheet's live DrawingsPart + root, or an ArgumentException when the
    // sheet has no drawing at all (a foreign/stale handle can't belong here).
    private (DrawingsPart Part, XDR.WorksheetDrawing Root) RequireDrawing(string paramName)
    {
        var drawingEl = Worksheet.GetFirstChild<S.Drawing>();
        if (drawingEl?.Id?.Value is string rid
            && _worksheetPart.GetPartById(rid) is DrawingsPart dp
            && dp.WorksheetDrawing is { } root)
            return (dp, root);
        throw StaleHandle(paramName, "drawing object");
    }

    private static ArgumentException StaleHandle(string paramName, string noun)
        => new($"This {noun} does not belong to this sheet (its anchor was not found — " +
               "a foreign handle, or one already removed).", paramName);

    // The image part a picture's blip (a:blip/@r:embed) resolves to, or null
    // when the picture carries no resolvable embed.
    private static ImagePart? ResolveBlipImagePart(DrawingsPart drawingsPart, XDR.Picture pic)
    {
        if (pic.BlipFill?.GetFirstChild<A.Blip>()?.Embed?.Value is { } embed
            && drawingsPart.TryGetPartById(embed, out var part) && part is ImagePart ip)
            return ip;
        return null;
    }

    // True when any surviving anchor's blip still resolves to imagePart
    // (ReferenceEquals on the part, so two distinct rel ids that point at one
    // shared part both keep it alive).
    private static bool AnyBlipReferences(XDR.WorksheetDrawing root, DrawingsPart drawingsPart, ImagePart imagePart)
        => root.Descendants<A.Blip>()
            .Select(b => b.Embed?.Value)
            .Where(embed => embed is not null)
            .Any(embed => drawingsPart.TryGetPartById(embed!, out var part) && ReferenceEquals(part, imagePart));

    // True when the anchor's graphicFrame references chartPart via its
    // c:chartReference r:id (the rel still resolves while the part is live).
    private static bool AnchorReferencesChart(DrawingsPart drawingsPart, OpenXmlElement anchor, ChartPart chartPart)
    {
        var rid = anchor.Descendants<XDR.GraphicFrame>().FirstOrDefault()
            ?.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>()?.Id?.Value;
        return rid is not null
            && drawingsPart.TryGetPartById(rid, out var part)
            && ReferenceEquals(part, chartPart);
    }

    // Drops the DrawingsPart and the <drawing> element wiring it once the last
    // anchor leaves — no empty drawing artifact (SDK-quirk #7 family). Verifies
    // the <drawing> rel still resolves to this part before removing the element,
    // mirroring DeleteVmlDrawingPart's guard.
    private void TeardownDrawingIfEmpty(DrawingsPart drawingsPart, XDR.WorksheetDrawing root)
    {
        if (root.ChildElements.Any(IsAnchor)) return;

        var drawingEl = Worksheet.GetFirstChild<S.Drawing>();
        if (drawingEl?.Id?.Value is { } rid
            && _worksheetPart.GetPartById(rid) is DrawingsPart wired
            && ReferenceEquals(wired, drawingsPart))
            drawingEl.Remove();
        _worksheetPart.DeletePart(drawingsPart);
    }
}
