// I-82 engine swap — cell comments (formulas/comments/hyperlinks slice)
// conformance.
//
// Mirrors the NPOI-engine CommentAndHyperlinkTests comment half on the Open XML
// SDK engine: the I11 default author, explicit authors, mutate-in-place replace
// semantics, chaining, null rejection, and the Save/Open round-trip. Adds
// part-graph assertions the NPOI tests could not express: the comments part +
// VML drawing part + <legacyDrawing> wiring that makes Excel actually SHOW the
// popup, author dedup, and one-VML-shape-per-comment (replace must not leak a
// second shape).

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class CommentTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-comment-{Guid.NewGuid():N}.xlsx");

    private static WorksheetPart SheetPart(IWorkbook wb)
        => wb.Underlying.WorkbookPart!.WorksheetParts.Single();

    // ---- Author semantics (decision I11) -------------------------------------

    [Fact]
    public void Comment_Default_Author_Is_NetXlsx_Per_I11()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Comment("hello there");

        sheet["A1"].GetComment().Should().Be("hello there");
        sheet["A1"].GetCommentAuthor().Should().Be("NetXlsx");
    }

    [Fact]
    public void Comment_Explicit_Author_Is_Used()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Comment("note", author: "qa-bot");
        sheet["A1"].GetCommentAuthor().Should().Be("qa-bot");
    }

    [Fact]
    public void Same_Author_On_Multiple_Comments_Is_Deduplicated()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Comment("one", "alice");
        sheet["B2"].Comment("two", "alice");
        sheet["C3"].Comment("three", "bob");

        var authors = SheetPart(wb).GetPartsOfType<WorksheetCommentsPart>().Single()
            .Comments!.Authors!.Elements<S.Author>().Select(a => a.Text).ToList();
        authors.Should().Equal("alice", "bob");
    }

    // ---- Replace semantics ----------------------------------------------------

    [Fact]
    public void Comment_Replaces_Existing_Comment_On_Same_Cell()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Comment("first");
        sheet["A1"].Comment("second", author: "alice");

        sheet["A1"].GetComment().Should().Be("second");
        sheet["A1"].GetCommentAuthor().Should().Be("alice");
    }

    [Fact]
    public void Replacing_A_Comment_Does_Not_Leak_A_Second_VML_Shape()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Comment("first");
        sheet["A1"].Comment("second");

        var vml = SheetPart(wb).GetPartsOfType<VmlDrawingPart>().Single();
        using var stream = vml.GetStream(FileMode.Open, FileAccess.Read);
        var doc = XDocument.Load(stream);
        XNamespace v = "urn:schemas-microsoft-com:vml";
        doc.Root!.Elements(v + "shape").Should().HaveCount(1,
            "mutate-in-place must reuse the existing popup shape");
    }

    // ---- Contract edges ---------------------------------------------------------

    [Fact]
    public void Comment_Returns_Same_Cell_For_Chaining()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var cell = sheet["A1"];
        cell.Comment("x").Should().BeSameAs(cell);
    }

    [Fact]
    public void Comment_Null_Text_Throws_ArgumentNullException()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].Comment(null!))).Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void GetComment_Returns_Null_When_No_Comment_Attached()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].GetComment().Should().BeNull();
        sheet["A1"].GetCommentAuthor().Should().BeNull();
    }

    // ---- Part graph (what makes Excel SHOW the popup) ---------------------------

    [Fact]
    public void Comment_Wires_CommentsPart_VmlPart_And_LegacyDrawing()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["B2"].Comment("reviewer flagged");

        var wsPart = SheetPart(wb);
        var commentsPart = wsPart.GetPartsOfType<WorksheetCommentsPart>().Single();
        var comment = commentsPart.Comments!.CommentList!.Elements<S.Comment>().Single();
        comment.Reference!.Value.Should().Be("B2");

        var legacy = wsPart.Worksheet!.GetFirstChild<S.LegacyDrawing>();
        legacy.Should().NotBeNull("Excel will not show a comment popup without the VML legacyDrawing");
        wsPart.GetPartById(legacy!.Id!.Value!).Should().BeOfType<VmlDrawingPart>();

        // The VML shape carries the Note client data anchored at the cell.
        using var stream = wsPart.GetPartsOfType<VmlDrawingPart>().Single()
            .GetStream(FileMode.Open, FileAccess.Read);
        var doc = XDocument.Load(stream);
        XNamespace v = "urn:schemas-microsoft-com:vml";
        XNamespace x = "urn:schemas-microsoft-com:office:excel";
        var shape = doc.Root!.Elements(v + "shape").Single();
        var client = shape.Element(x + "ClientData")!;
        ((string?)client.Attribute("ObjectType")).Should().Be("Note");
        ((string?)client.Element(x + "Row")).Should().Be("1");
        ((string?)client.Element(x + "Column")).Should().Be("1");
        ((string?)client.Element(x + "Anchor")).Should().Be("1, 0, 1, 0, 3, 0, 4, 0");
    }

    [Fact]
    public void Two_Comments_Share_One_VML_Part_With_Distinct_Shape_Ids()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["B2"].Comment("one");
        sheet["C3"].Comment("two");

        var vmlParts = SheetPart(wb).GetPartsOfType<VmlDrawingPart>().ToList();
        vmlParts.Should().HaveCount(1);
        using var stream = vmlParts[0].GetStream(FileMode.Open, FileAccess.Read);
        var doc = XDocument.Load(stream);
        XNamespace v = "urn:schemas-microsoft-com:vml";
        var ids = doc.Root!.Elements(v + "shape")
            .Select(s => (string?)s.Attribute("id")).ToList();
        ids.Should().HaveCount(2);
        ids.Should().OnlyHaveUniqueItems();
    }

    // ---- Round-trip ---------------------------------------------------------------

    [Fact]
    public void Comment_Roundtrips_Through_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("S");
                sheet["B2"].SetString("data");
                sheet["B2"].Comment("reviewer flagged");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var cell = wb["S"]["B2"];
                cell.GetComment().Should().Be("reviewer flagged");
                cell.GetCommentAuthor().Should().Be("NetXlsx");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Adding_A_Comment_To_An_Opened_File_Extends_The_Existing_Parts()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S")["B2"].Comment("first", "alice");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["S"]["D4"].Comment("second", "alice");

                wb["S"]["B2"].GetComment().Should().Be("first");
                wb["S"]["D4"].GetComment().Should().Be("second");

                // Still one comments part, one VML part, two shapes, one author.
                var wsPart = SheetPart(wb);
                wsPart.GetPartsOfType<WorksheetCommentsPart>().Should().HaveCount(1);
                var vmlParts = wsPart.GetPartsOfType<VmlDrawingPart>().ToList();
                vmlParts.Should().HaveCount(1);
                using var stream = vmlParts[0].GetStream(FileMode.Open, FileAccess.Read);
                XNamespace v = "urn:schemas-microsoft-com:vml";
                XDocument.Load(stream).Root!.Elements(v + "shape").Should().HaveCount(2);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
