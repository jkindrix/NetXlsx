// I-82 engine swap — Open XML SDK-backed cell comments (formulas/comments/
// hyperlinks slice).
//
// Mirrors the NPOI engine's ICell.Comment / GetComment / GetCommentAuthor
// contract (decision I11 — default author "NetXlsx"). An OOXML comment is TWO
// parts wired to the worksheet:
//   1. A WorksheetCommentsPart (xl/comments1.xml) — the <authors> list plus a
//      <comment ref> per cell carrying the text. This is the data.
//   2. A VmlDrawingPart (xl/drawings/vmlDrawing1.vml) referenced by a
//      <legacyDrawing r:id> worksheet child — the legacy VML shape that IS the
//      yellow popup box. Excel will not SHOW a comment without it.
// <legacyDrawing> is a strict-sequence CT_Worksheet child placed by
// OoxmlSchemaOrder (SDK-quirk #8). The VML blob is raw XML (an <xml> island,
// not OOXML-typed), so it is read/written through the part stream per
// SDK-quirk #12 — manipulated as an XDocument, never a typed DOM.
//
// Re-commenting a commented cell mutates the comment in place (text + author)
// and leaves its VML shape untouched — NPOI parity. Unlike NPOI, no empty
// zero-index author is emitted (NPOI's <author></author> artifact) and no empty
// xdr drawing part is created (NPOI's CreateDrawingPatriarch side effect);
// authorIds simply index the real authors. Both divergences are
// conformance-positive and invisible to the public surface.

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    private static readonly XNamespace VmlNs = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace OfficeNs = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace ExcelNs = "urn:schemas-microsoft-com:office:excel";

    // Applies (or replaces) the comment on one cell.
    internal void SetComment(int row, int col, string text, string? author)
    {
        var reference = CellAddress.Format(row, col);
        var authorName = author ?? "NetXlsx";   // Decision I11.

        var comments = GetOrCreateCommentsPart().Comments!;
        uint authorId = GetOrAddAuthor(comments.Authors!, authorName);

        var commentText = new S.CommentText(
            new S.Text(XStringCodec.Encode(text)) { Space = SpaceProcessingModeValues.Preserve });

        var existing = comments.CommentList!.Elements<S.Comment>()
            .FirstOrDefault(c => c.Reference?.Value == reference);
        if (existing is not null)
        {
            // Mutate in place — the cell already has a VML popup shape.
            existing.AuthorId = authorId;
            existing.CommentText = commentText;
            return;
        }

        comments.CommentList!.AppendChild(new S.Comment
        {
            Reference = reference,
            AuthorId = authorId,
            CommentText = commentText,
        });
        AddVmlCommentShape(row, col);
    }

    // The cell's comment text, or null when none is attached. Works for plain
    // and rich-text comments alike (InnerText concatenates the runs), matching
    // the NPOI engine's CellComment.String.String.
    internal string? GetComment(int row, int col)
        => FindComment(row, col)?.CommentText?.InnerText is { } raw ? XStringCodec.Decode(raw) : null;

    // The comment's author, or null when the cell has no comment. A corrupt
    // @authorId (non-integer, missing, or out of range of <authors>) fails loud
    // (decision I-83): OOXML defines no fallback for it. NPOI throws on
    // out-of-range too; for a non-integer id NPOI's deserializer silently
    // substitutes author 0 — exactly the silent-default trap I-83 forbids, so
    // the SDK engine is deliberately stricter there (pinned in the malformed
    // harness).
    internal string? GetCommentAuthor(int row, int col)
    {
        var comment = FindComment(row, col);
        if (comment is null) return null;
        var reference = CellAddress.Format(row, col);

        uint id;
        try
        {
            id = comment.AuthorId?.Value
                ?? throw new MalformedFileException($"cell {reference}: comment has no authorId");
        }
        catch (FormatException ex)
        {
            throw new MalformedFileException(
                $"cell {reference}: comment authorId '{comment.AuthorId?.InnerText}' is not a non-negative integer", ex);
        }

        var authors = CommentsPart?.Comments?.Authors?.Elements<S.Author>().ToList();
        if (authors is null || id >= authors.Count)
            throw new MalformedFileException(
                $"cell {reference}: comment authorId {id} is out of range of the authors list");
        return authors[(int)id].Text;
    }

    // ---- removal (I-91 removal family) --------------------------------------

    // Removes the cell's comment: the <comment> entry (with orphaned-author
    // bookkeeping), the legacy VML popup v:shape, and — when it was the last
    // comment on the sheet — the comments part entirely (no empty zero-index
    // artifact, mirroring the add path's no-empty-author discipline).
    // Idempotent: a no-op when the cell carries no comment.
    internal void RemoveComment(int row, int col)
    {
        var commentsPart = CommentsPart;
        var comments = commentsPart?.Comments;
        if (comments is null) return;
        var comment = FindComment(row, col);
        if (comment is null) return;

        uint? removedAuthorId = comment.AuthorId?.Value;
        comment.Remove();
        if (removedAuthorId is uint authorId)
            PruneOrphanAuthor(comments, authorId);

        bool lastComment = comments.CommentList?.Elements<S.Comment>().Any() != true;

        RemoveVmlCommentShape(row, col, deletePartWhenCommentsGone: lastComment);

        // No empty zero-index artifact: the whole comments part goes with the
        // last comment.
        if (lastComment) _worksheetPart.DeletePart(commentsPart!);
    }

    // Removes an author the just-removed comment was the last user of, then
    // re-indexes the survivors' authorIds — the inverse of GetOrAddAuthor, so
    // no stale author entry lingers. A single removal can orphan at most one
    // author.
    private static void PruneOrphanAuthor(S.Comments comments, uint authorId)
    {
        var list = comments.CommentList;
        if (list is null) return;
        if (list.Elements<S.Comment>().Any(c => c.AuthorId?.Value == authorId))
            return; // another comment still references this author

        var author = comments.Authors?.Elements<S.Author>().ElementAtOrDefault((int)authorId);
        if (author is null) return;
        author.Remove();

        // authorId is a 0-based index into <authors>; every survivor above the
        // removed slot shifts down by one.
        foreach (var c in list.Elements<S.Comment>())
            if (c.AuthorId?.Value is uint id && id > authorId)
                c.AuthorId = id - 1;
    }

    // Removes the cell's legacy VML popup v:shape and, when the last comment is
    // gone, the VML part itself — but only when that part holds no non-comment
    // shapes. [A-2026-06-11] VML safety guard: opened third-party files keep
    // form controls and other shapes in the same legacy VML part, and a
    // wholesale delete destroys them (PhpSpreadsheet #4105, ClosedXML #1285).
    private void RemoveVmlCommentShape(int row, int col, bool deletePartWhenCommentsGone)
    {
        var vmlPart = LocateVmlDrawingPart();
        if (vmlPart is null) return;

        XDocument doc;
        using (var read = vmlPart.GetStream(FileMode.Open, FileAccess.Read))
            doc = XDocument.Load(read);

        int r = row - 1, c = col - 1;
        doc.Root?.Elements(VmlNs + "shape")
            .FirstOrDefault(sh => IsCommentShapeFor(sh, r, c))
            ?.Remove();

        bool hasNonCommentShape = doc.Root?.Elements(VmlNs + "shape").Any(sh => !IsNoteShape(sh)) ?? false;

        if (deletePartWhenCommentsGone && !hasNonCommentShape)
        {
            DeleteVmlDrawingPart(vmlPart);
            return;
        }

        using var write = vmlPart.GetStream(FileMode.Create, FileAccess.Write);
        doc.Save(write);
    }

    // The VML part this sheet's comments use: prefer the part the live
    // <legacyDrawing r:id> wires (mirrors GetOrCreateVmlDrawingPart), falling
    // back to the part collection.
    private VmlDrawingPart? LocateVmlDrawingPart()
    {
        var legacy = Worksheet.GetFirstChild<S.LegacyDrawing>();
        if (legacy?.Id?.Value is { } relId
            && _worksheetPart.GetPartById(relId) is VmlDrawingPart wired)
            return wired;
        return _worksheetPart.GetPartsOfType<VmlDrawingPart>().FirstOrDefault();
    }

    // Drops the VML part and the <legacyDrawing> element wiring it, so no
    // dangling r:id is left behind.
    private void DeleteVmlDrawingPart(VmlDrawingPart vmlPart)
    {
        var legacy = Worksheet.GetFirstChild<S.LegacyDrawing>();
        if (legacy?.Id?.Value is { } relId
            && _worksheetPart.GetPartById(relId) is VmlDrawingPart wired
            && ReferenceEquals(wired, vmlPart))
            legacy.Remove();
        _worksheetPart.DeletePart(vmlPart);
    }

    private static bool IsCommentShapeFor(XElement shape, int r, int c)
    {
        var cd = shape.Element(ExcelNs + "ClientData");
        if (cd is null || (string?)cd.Attribute("ObjectType") != "Note") return false;
        return (string?)cd.Element(ExcelNs + "Row") == r.ToString(CultureInfo.InvariantCulture)
            && (string?)cd.Element(ExcelNs + "Column") == c.ToString(CultureInfo.InvariantCulture);
    }

    // A comment popup carries <x:ClientData ObjectType="Note">; anything else
    // (form controls, drop-downs, drawings) is a non-comment shape the guard
    // protects.
    private static bool IsNoteShape(XElement shape)
        => (string?)shape.Element(ExcelNs + "ClientData")?.Attribute("ObjectType") == "Note";

    // ---- comments part -------------------------------------------------------

    private WorksheetCommentsPart? CommentsPart
        => _worksheetPart.GetPartsOfType<WorksheetCommentsPart>().FirstOrDefault();

    private WorksheetCommentsPart GetOrCreateCommentsPart()
    {
        var part = CommentsPart;
        if (part is null)
        {
            part = _worksheetPart.AddNewPart<WorksheetCommentsPart>();
            part.Comments = new S.Comments(new S.Authors(), new S.CommentList());
        }
        return part;
    }

    private S.Comment? FindComment(int row, int col)
    {
        var reference = CellAddress.Format(row, col);
        return CommentsPart?.Comments?.CommentList?.Elements<S.Comment>()
            .FirstOrDefault(c => c.Reference?.Value == reference);
    }

    private static uint GetOrAddAuthor(S.Authors authors, string authorName)
    {
        uint index = 0;
        foreach (var a in authors.Elements<S.Author>())
        {
            if (a.Text == authorName) return index;
            index++;
        }
        authors.AppendChild(new S.Author(authorName));
        return index;
    }

    // ---- VML popup shape -------------------------------------------------------

    // Appends the legacy VML shape that renders the comment popup, creating the
    // VmlDrawingPart + <legacyDrawing r:id> wiring on first use. The shape
    // geometry mirrors the NPOI engine exactly: a hidden #_x0000_t202 note box
    // anchored 2 columns wide / 3 rows tall from the cell.
    private void AddVmlCommentShape(int row, int col)
    {
        var vmlPart = GetOrCreateVmlDrawingPart(out bool fresh);

        XDocument doc;
        if (fresh)
        {
            doc = new XDocument(VmlSkeleton());
        }
        else
        {
            using var read = vmlPart.GetStream(FileMode.Open, FileAccess.Read);
            doc = XDocument.Load(read);
            // An opened file's VML may carry only form controls etc.; make sure
            // the comment-box shapetype exists before referencing it.
            if (!doc.Root!.Elements(VmlNs + "shapetype")
                    .Any(st => (string?)st.Attribute("id") == "_x0000_t202"))
                doc.Root.Add(CommentShapeType());
        }

        doc.Root!.Add(CommentShape(NextVmlShapeId(doc), row, col));

        using var write = vmlPart.GetStream(FileMode.Create, FileAccess.Write);
        doc.Save(write);
    }

    private VmlDrawingPart GetOrCreateVmlDrawingPart(out bool fresh)
    {
        // An existing <legacyDrawing> wins: comments must join the VML part the
        // sheet already shows (form controls, prior comments).
        var legacy = Worksheet.GetFirstChild<S.LegacyDrawing>();
        if (legacy?.Id?.Value is { } relId
            && _worksheetPart.GetPartById(relId) is VmlDrawingPart wired)
        {
            fresh = false;
            return wired;
        }

        var part = _worksheetPart.GetPartsOfType<VmlDrawingPart>().FirstOrDefault();
        fresh = part is null;
        part ??= _worksheetPart.AddNewPart<VmlDrawingPart>();
        if (legacy is null)
        {
            var id = _worksheetPart.GetIdOfPart(part);
            OoxmlSchemaOrder.GetOrInsert(Worksheet, () => new S.LegacyDrawing { Id = id });
        }
        return part;
    }

    private static XElement VmlSkeleton()
        => new(XName.Get("xml"),
            new XAttribute(XNamespace.Xmlns + "v", VmlNs.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "o", OfficeNs.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "x", ExcelNs.NamespaceName),
            new XElement(OfficeNs + "shapelayout",
                new XAttribute(VmlNs + "ext", "edit"),
                new XElement(OfficeNs + "idmap",
                    new XAttribute(VmlNs + "ext", "edit"),
                    new XAttribute("data", "1"))),
            CommentShapeType());

    private static XElement CommentShapeType()
        => new(VmlNs + "shapetype",
            new XAttribute("coordsize", "21600,21600"),
            new XAttribute(OfficeNs + "spt", "202"),
            new XAttribute("id", "_x0000_t202"),
            new XAttribute("path", "m,l,21600r21600,l21600,xe"),
            new XElement(VmlNs + "stroke", new XAttribute("joinstyle", "miter")),
            new XElement(VmlNs + "path",
                new XAttribute(OfficeNs + "connecttype", "rect"),
                new XAttribute("gradientshapeok", "t")));

    private static XElement CommentShape(int shapeId, int row, int col)
    {
        // 0-based cell coordinates; the anchor spans (col..col+2, row..row+3)
        // with zero pixel offsets — the NPOI engine's 2-column-wide,
        // 3-row-tall popup.
        int r = row - 1, c = col - 1;
        var anchor = string.Create(CultureInfo.InvariantCulture,
            $"{c}, 0, {r}, 0, {c + 2}, 0, {r + 3}, 0");
        return new XElement(VmlNs + "shape",
            new XAttribute("id", string.Create(CultureInfo.InvariantCulture, $"_x0000_s{shapeId}")),
            new XAttribute("type", "#_x0000_t202"),
            new XAttribute("style", "position:absolute; visibility:hidden"),
            new XAttribute("fillcolor", "#ffffe1"),
            new XAttribute(OfficeNs + "insetmode", "auto"),
            new XElement(VmlNs + "fill",
                new XAttribute("type", "solid"),
                new XAttribute("color", "#ffffe1")),
            new XElement(VmlNs + "shadow",
                new XAttribute("on", "t"),
                new XAttribute("type", "single"),
                new XAttribute("obscured", "t"),
                new XAttribute("color", "black")),
            new XElement(VmlNs + "path", new XAttribute(OfficeNs + "connecttype", "none")),
            new XElement(VmlNs + "textbox", new XAttribute("style", "mso-direction-alt:auto")),
            new XElement(ExcelNs + "ClientData",
                new XAttribute("ObjectType", "Note"),
                new XElement(ExcelNs + "MoveWithCells"),
                new XElement(ExcelNs + "SizeWithCells"),
                new XElement(ExcelNs + "Anchor", anchor),
                new XElement(ExcelNs + "AutoFill", "false"),
                new XElement(ExcelNs + "Row", r.ToString(CultureInfo.InvariantCulture)),
                new XElement(ExcelNs + "Column", c.ToString(CultureInfo.InvariantCulture))));
    }

    // Next free shape id: NPOI numbers comment shapes _x0000_s1025, 1026, …;
    // scan whatever the part already carries so opened files never collide.
    private static int NextVmlShapeId(XDocument doc)
    {
        int max = 1024;
        foreach (var shape in doc.Root!.Elements(VmlNs + "shape"))
        {
            var id = (string?)shape.Attribute("id");
            if (id is not null && id.StartsWith("_x0000_s", StringComparison.Ordinal)
                && int.TryParse(id.AsSpan("_x0000_s".Length), NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out int n)
                && n > max)
                max = n;
        }
        return max + 1;
    }
}
