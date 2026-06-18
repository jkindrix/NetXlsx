// I-91 removal family, slice 1 of 2 (ledger R-11; folds in R-10): the simple
// element + rel half — ICell.RemoveHyperlink / RemoveComment,
// ISheet.RemoveValidation, IWorkbook.RemoveNamedRange, plus the OoxmlTable
// stale-handle retrofit. The drawing layer (pictures/charts/shapes/connectors,
// shared-media refcounting) is slice 2.
//
// These tests have no InternalsVisibleTo: every assertion is through the public
// API or the saved OOXML bytes (the real contract — engine-agnostic). Cleanup
// is checked at the zip/part level (element gone, part/rel gone when last,
// container dropped when empty), no-op-vs-throw is pinned per the design table,
// and the two [A-2026-06-11] amendments — the VML non-comment-shape guard and
// the x14 dual-stored validations — get their own opened-file fixtures.

using System;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xne = DocumentFormat.OpenXml.Office.Excel;

namespace NetXlsx.Tests;

public class RemovalFamilyTests
{
    private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace XExcel = "urn:schemas-microsoft-com:office:excel";

    // ---- shared inspection helpers --------------------------------------

    private static byte[] SavedBytes(IWorkbook wb)
    {
        using var ms = new MemoryStream();
        wb.Save(ms);
        return ms.ToArray();
    }

    private static System.Collections.Generic.List<string> ZipNames(byte[] bytes)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        return zip.Entries.Select(e => e.FullName).ToList();
    }

    // Counts package relationships of a given type across all parts (used to
    // assert the external-hyperlink reference relationship is gone).
    private static int RelationshipCount(byte[] bytes, string typeSuffix)
    {
        using var pkg = Package.Open(new MemoryStream(bytes, writable: false), FileMode.Open, FileAccess.Read);
        return pkg.GetParts()
            // .rels parts are infrastructure and cannot themselves hold relationships.
            .Where(p => p.ContentType != "application/vnd.openxmlformats-package.relationships+xml")
            .SelectMany(p => p.GetRelationships())
            .Count(r => r.RelationshipType.EndsWith(typeSuffix, StringComparison.Ordinal));
    }

    // Every internal relationship resolves to an existing part, and every
    // content part is reachable from the relationship graph (no orphans). Same
    // invariant SheetLifecycleTests pins for RemoveSheet.
    private static void AssertNoOrphanPartsOrRels(byte[] bytes)
    {
        using (var ms = new MemoryStream(bytes, writable: false))
        using (var pkg = Package.Open(ms, FileMode.Open, FileAccess.Read))
        {
            foreach (var part in pkg.GetParts())
            {
                if (part.ContentType == "application/vnd.openxmlformats-package.relationships+xml")
                    continue;
                foreach (var rel in part.GetRelationships())
                {
                    if (rel.TargetMode != TargetMode.Internal) continue;
                    var target = PackUriHelper.ResolvePartUri(part.Uri, rel.TargetUri);
                    pkg.PartExists(target).Should().BeTrue(
                        $"{part.Uri} relationship {rel.Id} → {rel.TargetUri} must resolve to an existing part");
                }
            }
        }

        using (var ms = new MemoryStream(bytes, writable: false))
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var reachable = new System.Collections.Generic.HashSet<string>(
                doc.GetAllParts().Select(p => p.Uri.ToString()), StringComparer.OrdinalIgnoreCase);
            foreach (var name in ZipNames(bytes))
            {
                if (name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)) continue;
                if (name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
                reachable.Should().Contain("/" + name,
                    $"saved part '{name}' must be reachable from the relationship graph (no orphan)");
            }
        }
    }

    private static XElement SheetRoot(IWorkbook wb, int sheetNumber = 1)
        => SavedOoxml.SheetXml(wb, sheetNumber).Root!;

    // ==================================================================
    //  RemoveHyperlink
    // ==================================================================

    [Fact]
    public void RemoveHyperlink_External_Drops_Element_And_Relationship_And_Container()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].Hyperlink("https://example.com/page", display: "link");

        // sanity: present before removal
        RelationshipCount(SavedBytes(wb), "/hyperlink").Should().Be(1);

        s["A1"].RemoveHyperlink();

        var root = SheetRoot(wb);
        root.Descendants(SavedOoxml.Main + "hyperlink").Should().BeEmpty("the element is gone");
        root.Element(SavedOoxml.Main + "hyperlinks").Should().BeNull("the empty container is dropped");
        var bytes = SavedBytes(wb);
        RelationshipCount(bytes, "/hyperlink").Should().Be(0, "the external reference relationship is gone");
        AssertNoOrphanPartsOrRels(bytes);
    }

    [Fact]
    public void RemoveHyperlink_Internal_Drops_Element_And_Container()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Target");
        var s = wb.AddSheet("S");
        s["A1"].Hyperlink("#Target!A1", display: "jump");

        s["A1"].RemoveHyperlink();

        var root = SheetRoot(wb, 2);
        root.Descendants(SavedOoxml.Main + "hyperlink").Should().BeEmpty();
        root.Element(SavedOoxml.Main + "hyperlinks").Should().BeNull();
        s["A1"].GetHyperlink().Should().BeNull();
    }

    [Fact]
    public void RemoveHyperlink_Is_Idempotent_NoOp_When_Absent()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");

        Action removeUnset = () => s["A1"].RemoveHyperlink();
        removeUnset.Should().NotThrow("absent hyperlink removal is a no-op");

        s["A1"].Hyperlink("https://example.com");
        s["A1"].RemoveHyperlink();
        Action removeAgain = () => s["A1"].RemoveHyperlink();
        removeAgain.Should().NotThrow("a second removal is also a no-op");
    }

    [Fact]
    public void RemoveHyperlink_Keeps_Other_Links_And_Container()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].Hyperlink("https://a.example", display: "a");
        s["A2"].Hyperlink("https://b.example", display: "b");

        s["A1"].RemoveHyperlink();

        var root = SheetRoot(wb);
        var links = root.Element(SavedOoxml.Main + "hyperlinks");
        links.Should().NotBeNull("a surviving link keeps the container");
        links!.Elements(SavedOoxml.Main + "hyperlink").Should().HaveCount(1);
        s["A2"].GetHyperlink().Should().Be("https://b.example");
        RelationshipCount(SavedBytes(wb), "/hyperlink").Should().Be(1);
    }

    [Fact]
    public void RemoveHyperlink_Leaves_Display_Text()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].Hyperlink("https://example.com", display: "click me");

        s["A1"].RemoveHyperlink();

        s["A1"].GetString().Should().Be("click me", "removal detaches the link, not the text");
        s["A1"].GetHyperlink().Should().BeNull();
    }

    // R-10 regression: Hyperlink → RemoveHyperlink → SetString leaves no rel and
    // no <hyperlink> element (the defect was the missing exit + missing docs).
    [Fact]
    public void R10_Regression_RemoveHyperlink_Then_SetString_Leaves_No_Rel_Or_Element()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].Hyperlink("https://old.example/url", display: "old");
        s["A1"].RemoveHyperlink();
        s["A1"].SetString("plain text now");

        var bytes = SavedBytes(wb);
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var sheet = XDocument.Load(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
        sheet.Descendants(SavedOoxml.Main + "hyperlink").Should().BeEmpty("no <hyperlink> element survives");
        RelationshipCount(bytes, "/hyperlink").Should().Be(0, "no dangling external relationship survives");
        s["A1"].GetString().Should().Be("plain text now");
        AssertNoOrphanPartsOrRels(bytes);
    }

    // ==================================================================
    //  RemoveComment
    // ==================================================================

    [Fact]
    public void RemoveComment_Last_Drops_Comments_And_Vml_Parts()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].Comment("a note", author: "qa");

        ZipNames(SavedBytes(wb)).Should().Contain(n => n.Contains("comments", StringComparison.OrdinalIgnoreCase));

        s["A1"].RemoveComment();

        var bytes = SavedBytes(wb);
        var names = ZipNames(bytes);
        names.Should().NotContain(n => n.Contains("comments", StringComparison.OrdinalIgnoreCase),
            "the last comment drops the comments part — no empty zero-index artifact");
        names.Should().NotContain(n => n.EndsWith(".vml", StringComparison.OrdinalIgnoreCase),
            "the VML part goes too when it held only comment shapes");
        SheetRoot(wb).Element(SavedOoxml.Main + "legacyDrawing").Should().BeNull(
            "the <legacyDrawing> wiring is removed with the VML part");
        s["A1"].GetComment().Should().BeNull();
        AssertNoOrphanPartsOrRels(bytes);
    }

    [Fact]
    public void RemoveComment_Is_Idempotent_NoOp_When_Absent()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");

        Action removeUnset = () => s["A1"].RemoveComment();
        removeUnset.Should().NotThrow();

        s["A1"].Comment("x");
        s["A1"].RemoveComment();
        Action removeAgain = () => s["A1"].RemoveComment();
        removeAgain.Should().NotThrow();
    }

    [Fact]
    public void RemoveComment_Keeps_Other_Comments_And_Prunes_Orphan_Author()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].Comment("first", author: "alice");
        s["A2"].Comment("second", author: "bob");

        s["A1"].RemoveComment();

        // The survivor's text AND author are intact after the orphaned author
        // ("alice") is pruned and authorIds re-indexed.
        s["A1"].GetComment().Should().BeNull();
        s["A2"].GetComment().Should().Be("second");
        s["A2"].GetCommentAuthor().Should().Be("bob");

        var bytes = SavedBytes(wb);
        ZipNames(bytes).Should().Contain(n => n.Contains("comments", StringComparison.OrdinalIgnoreCase),
            "other comments keep the part alive");
        AssertNoOrphanPartsOrRels(bytes);
    }

    [Fact]
    public void RemoveComment_Survives_Save_And_Reopen()
    {
        var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("S");
            s["A1"].Comment("one", author: "x");
            s["B2"].Comment("two", author: "y");
            s["A1"].RemoveComment();
            wb.Save(ms);
        }
        ms.Position = 0;
        using var reopened = Workbook.Open(ms, leaveOpen: false);
        var sheet = reopened["S"];
        sheet["A1"].GetComment().Should().BeNull();
        sheet["B2"].GetComment().Should().Be("two");
        sheet["B2"].GetCommentAuthor().Should().Be("y");
    }

    // [A-2026-06-11] VML safety guard: a legacy VML part that also carries a
    // non-comment shape (form control, etc. — common in opened third-party
    // files) must survive the removal of its last comment; only the comment's
    // v:shape goes. Crafted via the Underlying hatch to mimic such a file.
    [Fact]
    public void RemoveComment_Preserves_Vml_Part_With_NonComment_Shape()
    {
        var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("S");
            s["A1"].Comment("note");   // creates the VML part + comment shape

            // Inject a non-comment shape (a form-control button) into the same
            // legacy VML part through the part stream.
            var wsPart = wb.Underlying.WorkbookPart!.WorksheetParts.First();
            var vml = wsPart.GetPartsOfType<VmlDrawingPart>().First();
            XDocument doc;
            using (var read = vml.GetStream(FileMode.Open, FileAccess.Read))
                doc = XDocument.Load(read);
            doc.Root!.Add(new XElement(Vml + "shape",
                new XAttribute("id", "_x0000_s2050"),
                new XAttribute("type", "#_x0000_t201"),
                new XElement(XExcel + "ClientData",
                    new XAttribute("ObjectType", "Button"))));
            using (var write = vml.GetStream(FileMode.Create, FileAccess.Write))
                doc.Save(write);

            wb.Save(ms);
        }

        // Reopen the crafted file, remove the only comment.
        ms.Position = 0;
        var bytes2 = new MemoryStream();
        using (var wb = Workbook.Open(ms, leaveOpen: false))
        {
            wb["S"]["A1"].RemoveComment();
            wb.Save(bytes2);
        }

        var saved = bytes2.ToArray();
        var names = ZipNames(saved);
        names.Should().NotContain(n => n.Contains("comments", StringComparison.OrdinalIgnoreCase),
            "the comments part still goes with the last comment");
        names.Should().Contain(n => n.EndsWith(".vml", StringComparison.OrdinalIgnoreCase),
            "the VML part survives because it still holds a non-comment shape");

        // The button survives; the comment's note shape is gone.
        using var zip = new ZipArchive(new MemoryStream(saved), ZipArchiveMode.Read);
        var vmlEntry = zip.Entries.First(e => e.FullName.EndsWith(".vml", StringComparison.OrdinalIgnoreCase));
        var vdoc = XDocument.Load(vmlEntry.Open());
        var shapes = vdoc.Root!.Elements(Vml + "shape").ToList();
        shapes.Should().ContainSingle();
        ((string?)shapes[0].Element(XExcel + "ClientData")?.Attribute("ObjectType")).Should().Be("Button");
        AssertNoOrphanPartsOrRels(saved);
    }

    // ==================================================================
    //  RemoveValidation
    // ==================================================================

    [Fact]
    public void RemoveValidation_Removes_Rule_And_Drops_Container()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddValidation("A1:A5", DataValidation.IntegerBetween(1, 10));

        s.RemoveValidation("A1:A5");

        SheetRoot(wb).Element(SavedOoxml.Main + "dataValidations").Should().BeNull(
            "the empty container is dropped");
    }

    [Fact]
    public void RemoveValidation_Keeps_Others_And_Fixes_Count()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddValidation("A1:A5", DataValidation.IntegerBetween(1, 10));
        s.AddValidation("B1:B5", DataValidation.IntegerBetween(1, 10));

        s.RemoveValidation("A1:A5");

        var container = SheetRoot(wb).Element(SavedOoxml.Main + "dataValidations")!;
        container.Elements(SavedOoxml.Main + "dataValidation").Should().HaveCount(1);
        ((string?)container.Attribute("count")).Should().Be("1", "@count is kept in sync");
        var sqref = (string?)container.Elements(SavedOoxml.Main + "dataValidation").Single().Attribute("sqref");
        sqref.Should().Be("B1:B5");
    }

    [Fact]
    public void RemoveValidation_Matches_By_Canonical_Range()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddValidation("A2", DataValidation.IntegerBetween(1, 10));   // stored as "A2"

        // "A2:A2" is the same range, canonicalized — must match.
        s.RemoveValidation("A2:A2");

        SheetRoot(wb).Element(SavedOoxml.Main + "dataValidations").Should().BeNull();
    }

    [Fact]
    public void RemoveValidation_Throws_When_Absent()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddValidation("A1:A5", DataValidation.IntegerBetween(1, 10));

        Action act = () => s.RemoveValidation("Z1:Z5");
        act.Should().Throw<ArgumentException>("a missing exact range throws, not a silent no-op")
            .And.ParamName.Should().Be("a1Range");

        // The present rule is untouched.
        SheetRoot(wb).Descendants(SavedOoxml.Main + "dataValidation").Should().HaveCount(1);
    }

    [Fact]
    public void RemoveValidation_Throws_On_Null()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.RemoveValidation(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    // [A-2026-06-11] Dual storage: a cross-sheet-list-source validation lives in
    // x14:dataValidations inside <extLst>, keyed by <xm:sqref>. Crafted via
    // Underlying, saved/reopened (so it parses as an opened file would), then
    // removed — the emptied x14 container AND its <ext> wrapper must both go.
    [Fact]
    public void RemoveValidation_Removes_X14_DualStored_Rule_And_Ext()
    {
        var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            wb.AddSheet("Other");
            var s = wb.AddSheet("S");
            var ws = (DocumentFormat.OpenXml.Spreadsheet.Worksheet)s.Underlying;

            var x14dvs = new X14.DataValidations(
                new X14.DataValidation(
                    new X14.DataValidationForumla1(new Xne.Formula("Other!$A$1:$A$3")),
                    new Xne.ReferenceSequence("D4"))
                {
                    Type = S.DataValidationValues.List,
                    AllowBlank = true,
                });
            x14dvs.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");

            ws.AppendChild(new S.WorksheetExtensionList(
                new S.WorksheetExtension(x14dvs) { Uri = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" }));
            wb.Save(ms);
        }

        ms.Position = 0;
        var outBytes = new MemoryStream();
        using (var wb = Workbook.Open(ms, leaveOpen: false))
        {
            // sanity: the x14 rule round-tripped
            SheetRoot(wb, 2).Descendants(
                XName.Get("dataValidations", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"))
                .Should().ContainSingle("the crafted x14 container is present before removal");

            wb["S"].RemoveValidation("D4");
            wb.Save(outBytes);
        }

        using var zip = new ZipArchive(new MemoryStream(outBytes.ToArray()), ZipArchiveMode.Read);
        var sheet = XDocument.Load(zip.GetEntry("xl/worksheets/sheet2.xml")!.Open());
        XNamespace x14ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        sheet.Descendants(x14ns + "dataValidations").Should().BeEmpty("the x14 container is gone");
        sheet.Descendants(SavedOoxml.Main + "extLst").Should().BeEmpty(
            "the emptied <ext>/<extLst> are removed (an emptied-but-present ext triggers Excel repair)");
        AssertNoOrphanPartsOrRels(outBytes.ToArray());
    }

    // ==================================================================
    //  RemoveNamedRange
    // ==================================================================

    [Fact]
    public void RemoveNamedRange_Removes_Name_And_Drops_Container()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.AddNamedRange("Sales", "S!$A$1:$B$10");

        wb.RemoveNamedRange("Sales");

        wb.NamedRanges.Should().BeEmpty();
        SavedOoxml.WorkbookXml(wb).Root!
            .Element(SavedOoxml.Main + "definedNames").Should().BeNull("the empty container is dropped");
    }

    [Fact]
    public void RemoveNamedRange_Keeps_Others()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.AddNamedRange("Sales", "S!$A$1");
        wb.AddNamedRange("Costs", "S!$B$1");

        wb.RemoveNamedRange("Sales");

        wb.NamedRanges.Select(n => n.Name).Should().ContainSingle().Which.Should().Be("Costs");
    }

    [Fact]
    public void RemoveNamedRange_Is_Case_Insensitive()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.AddNamedRange("Sales", "S!$A$1");

        wb.RemoveNamedRange("SALES");   // names are unique case-insensitive (I-9)

        wb.NamedRanges.Should().BeEmpty();
    }

    [Fact]
    public void RemoveNamedRange_Throws_When_Absent()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.AddNamedRange("Sales", "S!$A$1");

        Action act = () => wb.RemoveNamedRange("Nope");
        act.Should().Throw<ArgumentException>().And.ParamName.Should().Be("name");

        wb.NamedRanges.Should().ContainSingle();
    }

    [Fact]
    public void RemoveNamedRange_Throws_On_Null()
    {
        using var wb = Workbook.Create();
        Action act = () => wb.RemoveNamedRange(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    // ==================================================================
    //  OoxmlTable stale-handle retrofit
    // ==================================================================

    [Fact]
    public void RemovedTable_Handle_Members_Throw_InvalidOperation()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("H1");
        s["A2"].SetNumber(1);
        var table = s.AddTable("A1:A2", "T");

        s.RemoveTable(table);

        // Distinct from the disposed-workbook ObjectDisposedException: a removed
        // handle surfaces our own InvalidOperationException (message pinned so
        // it's the explicit guard, not an incidental SDK deleted-part throw).
        table.Invoking(t => { var _ = t.Name; }).Should().Throw<InvalidOperationException>()
            .WithMessage("*removed*");
        table.Invoking(t => { var _ = t.Address; }).Should().Throw<InvalidOperationException>()
            .WithMessage("*removed*");
        table.Invoking(t => t.AddTotalsRow()).Should().Throw<InvalidOperationException>()
            .WithMessage("*removed*");
        table.Invoking(t => { var _ = t.ColumnNames; }).Should().Throw<InvalidOperationException>()
            .WithMessage("*removed*");
        table.Invoking(t => { var _ = t.DisplayName; }).Should().Throw<InvalidOperationException>()
            .WithMessage("*removed*");
        table.Invoking(t => { var _ = t.Sheet; }).Should().Throw<InvalidOperationException>()
            .WithMessage("*removed*");
        table.Invoking(t => { var _ = t.HasTotalsRow; }).Should().Throw<InvalidOperationException>()
            .WithMessage("*removed*");
        table.Invoking(t => { var _ = t.StyleName; }).Should().Throw<InvalidOperationException>()
            .WithMessage("*removed*");
        table.Invoking(t => { var _ = t.Underlying; }).Should().Throw<InvalidOperationException>()
            .WithMessage("*removed*");
    }

    [Fact]
    public void RemovedTable_Handle_After_Dispose_Throws_ObjectDisposed()
    {
        var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("H1");
        s["A2"].SetNumber(1);
        var table = s.AddTable("A1:A2", "T");
        wb.Dispose();

        // Disposal is checked first, so a live-but-disposed handle (never
        // removed) raises ObjectDisposedException.
        table.Invoking(t => { var _ = t.Name; }).Should().Throw<ObjectDisposedException>();
    }

    // ==================================================================
    //  Drawing layer (I-91 slice 2): RemovePicture / RemoveChart /
    //  RemoveShape / RemoveConnector — anchor + shared-media refcount +
    //  drawing-part teardown + foreign / stale / removed-handle semantics.
    // ==================================================================

    private static readonly XNamespace Xdr =
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    // 8-byte PNG signature: enough for the explicit-format embed path (the
    // engine writes the bytes verbatim and falls back to a 1x1 extent).
    private static readonly byte[] Png = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

    private static XDocument? DrawingXml(byte[] bytes, int n = 1)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var entry = zip.GetEntry($"xl/drawings/drawing{n}.xml");
        return entry is null ? null : XDocument.Load(entry.Open());
    }

    private static int AnchorCount(XDocument? drawing)
        => drawing is null ? 0
           : drawing.Root!.Elements()
               .Count(e => e.Name == Xdr + "twoCellAnchor" || e.Name == Xdr + "oneCellAnchor");

    private static int MediaPartCount(byte[] bytes)
        => ZipNames(bytes).Count(n => n.StartsWith("xl/media/", StringComparison.OrdinalIgnoreCase));

    // ---- RemovePicture --------------------------------------------------

    [Fact]
    public void RemovePicture_Last_Drops_Anchor_Image_Part_And_Drawing()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var pic = s.AddPicture("A1", Png, ImageFormat.Png);

        MediaPartCount(SavedBytes(wb)).Should().Be(1, "sanity: the image part is present before removal");

        s.RemovePicture(pic);

        var bytes = SavedBytes(wb);
        MediaPartCount(bytes).Should().Be(0, "the only image part goes with the last picture");
        ZipNames(bytes).Should().NotContain(n => n.Contains("xl/drawings/drawing", StringComparison.OrdinalIgnoreCase),
            "the empty drawing part is dropped");
        SheetRoot(wb).Element(SavedOoxml.Main + "drawing").Should().BeNull(
            "the worksheet <drawing> rel is removed with the part");
        AssertNoOrphanPartsOrRels(bytes);
    }

    [Fact]
    public void RemovePicture_Keeps_Other_Picture_And_Drawing()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var pic1 = s.AddPicture("A1", Png, ImageFormat.Png);
        s.AddPicture("C3", Png, ImageFormat.Png);

        // AddPicture does not dedup: two pictures => two distinct image parts.
        MediaPartCount(SavedBytes(wb)).Should().Be(2);

        s.RemovePicture(pic1);

        var bytes = SavedBytes(wb);
        AnchorCount(DrawingXml(bytes)).Should().Be(1, "the surviving picture keeps its anchor");
        MediaPartCount(bytes).Should().Be(1, "only the removed picture's (unshared) image part goes");
        s.Pictures.Should().ContainSingle();
        AssertNoOrphanPartsOrRels(bytes);
    }

    // The headline of this slice [A-2026-06-11]: two pictures sharing ONE image
    // part. AddPicture doesn't dedup, so the shared rel is crafted via the
    // Underlying hatch (point pic2's blip @r:embed at pic1's image rel, then
    // drop pic2's now-orphaned own part). Removing one keeps the part alive for
    // the other; removing both deletes it.
    [Fact]
    public void RemovePicture_RefCounts_Shared_Image_Part()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var pic1 = s.AddPicture("A1", Png, ImageFormat.Png);
        var pic2 = s.AddPicture("C3", Png, ImageFormat.Png);

        var dp = wb.Underlying.WorkbookPart!.WorksheetParts.First()
            .GetPartsOfType<DrawingsPart>().First();
        var blip1 = pic1.Underlying.BlipFill!.GetFirstChild<A.Blip>()!;
        var blip2 = pic2.Underlying.BlipFill!.GetFirstChild<A.Blip>()!;
        string pic1Embed = blip1.Embed!.Value!;
        var orphan = (ImagePart)dp.GetPartById(blip2.Embed!.Value!);
        blip2.Embed = pic1Embed;     // pic2 now shares pic1's image part
        dp.DeletePart(orphan);       // drop pic2's own (now-unreferenced) part

        MediaPartCount(SavedBytes(wb)).Should().Be(1, "sanity: the two pictures now share one image part");

        // Remove one consumer — the shared part SURVIVES, the other picture stays.
        s.RemovePicture(pic1);
        var afterOne = SavedBytes(wb);
        MediaPartCount(afterOne).Should().Be(1, "the image part survives while another anchor references it");
        AnchorCount(DrawingXml(afterOne)).Should().Be(1, "the second picture still renders");
        AssertNoOrphanPartsOrRels(afterOne);

        // Remove the last consumer — now the shared part GOES.
        s.RemovePicture(pic2);
        var afterBoth = SavedBytes(wb);
        MediaPartCount(afterBoth).Should().Be(0, "the shared image part goes once its last anchor leaves");
        ZipNames(afterBoth).Should().NotContain(n => n.Contains("xl/drawings/drawing", StringComparison.OrdinalIgnoreCase));
        AssertNoOrphanPartsOrRels(afterBoth);
    }

    [Fact]
    public void RemovePicture_Foreign_Handle_Throws()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddPicture("A1", Png, ImageFormat.Png);

        Action act = () => s.RemovePicture(new ForeignPicture());
        act.Should().Throw<ArgumentException>("a non-Ooxml picture is foreign")
            .And.ParamName.Should().Be("picture");
    }

    [Fact]
    public void RemovePicture_Stale_Handle_DoubleRemove_Throws()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var pic = s.AddPicture("A1", Png, ImageFormat.Png);
        s.RemovePicture(pic);

        Action act = () => s.RemovePicture(pic);
        act.Should().Throw<ArgumentException>("a second removal finds no anchor (stale)")
            .And.ParamName.Should().Be("picture");
    }

    [Fact]
    public void RemovePicture_Throws_On_Null()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.RemovePicture(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void RemovedPicture_Handle_Members_Throw_InvalidOperation()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var pic = s.AddPicture("A1", Png, ImageFormat.Png);

        s.RemovePicture(pic);

        // Our explicit guard (message pinned), distinct from the disposed-workbook
        // ObjectDisposedException and from an incidental SDK deleted-part throw.
        pic.Invoking(p => { var _ = p.Sheet; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        pic.Invoking(p => { var _ = p.Format; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        pic.Invoking(p => { var _ = p.FromCell; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        pic.Invoking(p => { var _ = p.Data; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        pic.Invoking(p => { var _ = p.Border; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        pic.Invoking(p => { var _ = p.Underlying; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
    }

    // ---- RemoveChart ----------------------------------------------------

    [Fact]
    public void RemoveChart_Last_Drops_Anchor_Chart_Part_And_Drawing()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var chart = s.AddChart(ChartType.Line, "D1", "K15", "A1:A5", "B1:B5");

        ZipNames(SavedBytes(wb)).Should().Contain(n => n.Contains("charts/chart", StringComparison.OrdinalIgnoreCase));

        s.RemoveChart(chart);

        var bytes = SavedBytes(wb);
        ZipNames(bytes).Should().NotContain(n => n.Contains("charts/", StringComparison.OrdinalIgnoreCase),
            "the chart part goes with its anchor (charts are not shared)");
        ZipNames(bytes).Should().NotContain(n => n.Contains("xl/drawings/drawing", StringComparison.OrdinalIgnoreCase),
            "the empty drawing part is dropped");
        SheetRoot(wb).Element(SavedOoxml.Main + "drawing").Should().BeNull();
        AssertNoOrphanPartsOrRels(bytes);
    }

    [Fact]
    public void RemoveChart_Foreign_Handle_Throws()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A5", "B1:B5");

        Action act = () => s.RemoveChart(new ForeignChart());
        act.Should().Throw<ArgumentException>().And.ParamName.Should().Be("chart");
    }

    [Fact]
    public void RemoveChart_Stale_Handle_DoubleRemove_Throws()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var chart = s.AddChart(ChartType.Line, "D1", "K15", "A1:A5", "B1:B5");
        s.RemoveChart(chart);

        Action act = () => s.RemoveChart(chart);
        act.Should().Throw<ArgumentException>().And.ParamName.Should().Be("chart");
    }

    [Fact]
    public void RemovedChart_Handle_Members_Throw_InvalidOperation()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var chart = s.AddChart(ChartType.Line, "D1", "K15", "A1:A5", "B1:B5");

        s.RemoveChart(chart);

        chart.Invoking(c => { var _ = c.Sheet; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        chart.Invoking(c => { var _ = c.Type; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        chart.Invoking(c => c.SetTitle("x")).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        chart.Invoking(c => { var _ = c.Underlying; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
    }

    // ---- RemoveShape ----------------------------------------------------

    [Fact]
    public void RemoveShape_Last_Drops_Anchor_And_Drawing()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var shape = s.AddShape(ShapeType.Rectangle, "A1", "C3");

        s.RemoveShape(shape);

        var bytes = SavedBytes(wb);
        ZipNames(bytes).Should().NotContain(n => n.Contains("xl/drawings/drawing", StringComparison.OrdinalIgnoreCase),
            "a shape owns no part; the empty drawing is dropped");
        SheetRoot(wb).Element(SavedOoxml.Main + "drawing").Should().BeNull();
        AssertNoOrphanPartsOrRels(bytes);
    }

    [Fact]
    public void RemoveShape_Keeps_Other_Shape()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var shape1 = s.AddShape(ShapeType.Rectangle, "A1", "B2");
        s.AddShape(ShapeType.Ellipse, "C3", "D4");

        s.RemoveShape(shape1);

        var bytes = SavedBytes(wb);
        AnchorCount(DrawingXml(bytes)).Should().Be(1, "the second shape keeps its anchor");
        AssertNoOrphanPartsOrRels(bytes);
    }

    [Fact]
    public void RemoveShape_Foreign_Handle_Throws()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddShape(ShapeType.Rectangle, "A1", "C3");

        Action act = () => s.RemoveShape(new ForeignShape());
        act.Should().Throw<ArgumentException>().And.ParamName.Should().Be("shape");
    }

    [Fact]
    public void RemovedShape_Handle_Members_Throw_InvalidOperation()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var shape = s.AddShape(ShapeType.Rectangle, "A1", "C3");

        s.RemoveShape(shape);

        shape.Invoking(sh => { var _ = sh.Sheet; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        shape.Invoking(sh => { var _ = sh.Type; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        shape.Invoking(sh => { var _ = sh.Underlying; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
    }

    // ---- RemoveConnector ------------------------------------------------

    [Fact]
    public void RemoveConnector_Last_Drops_Anchor_And_Drawing()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var connector = s.AddConnector(ConnectorType.Straight, "A1", "C3");

        s.RemoveConnector(connector);

        var bytes = SavedBytes(wb);
        ZipNames(bytes).Should().NotContain(n => n.Contains("xl/drawings/drawing", StringComparison.OrdinalIgnoreCase),
            "a connector owns no part; the empty drawing is dropped");
        SheetRoot(wb).Element(SavedOoxml.Main + "drawing").Should().BeNull();
        s.Connectors.Should().BeEmpty();
        AssertNoOrphanPartsOrRels(bytes);
    }

    [Fact]
    public void RemoveConnector_Keeps_Other_Connector()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var connector1 = s.AddConnector(ConnectorType.Straight, "A1", "B2");
        s.AddConnector(ConnectorType.Bent, "C3", "D4");

        s.RemoveConnector(connector1);

        var bytes = SavedBytes(wb);
        AnchorCount(DrawingXml(bytes)).Should().Be(1, "the second connector keeps its anchor");
        s.Connectors.Should().ContainSingle();
        AssertNoOrphanPartsOrRels(bytes);
    }

    [Fact]
    public void RemoveConnector_Foreign_Handle_Throws()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConnector(ConnectorType.Straight, "A1", "C3");

        Action act = () => s.RemoveConnector(new ForeignConnector());
        act.Should().Throw<ArgumentException>().And.ParamName.Should().Be("connector");
    }

    [Fact]
    public void RemovedConnector_Handle_Members_Throw_InvalidOperation()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var connector = s.AddConnector(ConnectorType.Straight, "A1", "C3");

        s.RemoveConnector(connector);

        connector.Invoking(c => { var _ = c.Sheet; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        connector.Invoking(c => { var _ = c.Type; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        connector.Invoking(c => { var _ = c.FromCell; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
        connector.Invoking(c => { var _ = c.Underlying; }).Should().Throw<InvalidOperationException>().WithMessage("*removed*");
    }

    // ---- Mixed: removing one drawing kind leaves the others -------------

    [Fact]
    public void RemovePicture_Leaves_Coexisting_Chart_And_Shape()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var pic = s.AddPicture("A1", Png, ImageFormat.Png);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A5", "B1:B5");
        s.AddShape(ShapeType.Rectangle, "M1", "N4");

        s.RemovePicture(pic);

        var bytes = SavedBytes(wb);
        AnchorCount(DrawingXml(bytes)).Should().Be(2, "the chart and shape anchors remain");
        MediaPartCount(bytes).Should().Be(0, "the picture's image part goes");
        ZipNames(bytes).Should().Contain(n => n.Contains("charts/chart", StringComparison.OrdinalIgnoreCase),
            "the chart part survives");
        AssertNoOrphanPartsOrRels(bytes);
    }
}
