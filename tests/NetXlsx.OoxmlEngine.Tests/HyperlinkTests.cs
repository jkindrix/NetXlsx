// I-82 engine swap — hyperlinks (formulas/comments/hyperlinks slice) conformance.
//
// Mirrors the NPOI-engine CommentAndHyperlinkTests hyperlink half on the Open
// XML SDK engine: the I13 scheme sniff (http(s)/mailto/file external, #internal
// location, everything else rejected), display-text semantics, verbatim target
// preservation (no URI canonicalization), replace semantics (incl. relationship
// cleanup), and the Save/Open round-trip.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class HyperlinkTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-hyperlink-{Guid.NewGuid():N}.xlsx");

    private static WorksheetPart SheetPart(IWorkbook wb)
        => wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.First();

    // ---- Scheme sniff (decision I13) ----------------------------------------

    [Theory]
    [InlineData("https://example.com")]
    [InlineData("http://example.com/path?q=1")]
    [InlineData("HTTPS://Example.COM")]                  // case-insensitive scheme
    [InlineData("mailto:user@example.com")]
    [InlineData("file:///c/temp/file.xlsx")]
    public void Hyperlink_Accepts_Supported_Schemes(string target)
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Hyperlink(target);
        sheet["A1"].GetHyperlink().Should().Be(target);
    }

    [Fact]
    public void Hyperlink_Internal_Document_Form_Strips_Leading_Hash()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Other");
        var sheet = wb.AddSheet("S");
        sheet["A1"].Hyperlink("#Other!A1");
        sheet["A1"].GetHyperlink().Should().Be("Other!A1");
    }

    [Theory]
    [InlineData("ftp://example.com")]
    [InlineData("javascript:alert(1)")]
    [InlineData("not-a-url")]
    [InlineData("/absolute/path")]
    public void Hyperlink_Rejects_Unsupported_Scheme(string target)
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].Hyperlink(target)))
            .Should().Throw<ArgumentException>()
            .WithMessage("*unsupported scheme*");
    }

    [Fact]
    public void Hyperlink_Null_Target_Throws_ArgumentNullException()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].Hyperlink(null!))).Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Hyperlink_Empty_Target_Throws_ArgumentException()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].Hyperlink(""))).Should().Throw<ArgumentException>();
    }

    // ---- Display-text semantics ----------------------------------------------

    [Fact]
    public void Hyperlink_With_Display_Sets_Cell_Text()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Hyperlink("https://example.com", display: "Example");
        sheet["A1"].GetString().Should().Be("Example");
        sheet["A1"].GetHyperlink().Should().Be("https://example.com");
    }

    [Fact]
    public void Hyperlink_Without_Display_On_Empty_Cell_Sets_Text_To_Target()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Hyperlink("https://example.com");
        sheet["A1"].GetString().Should().Be("https://example.com");
    }

    [Fact]
    public void Hyperlink_Without_Display_Preserves_Existing_Cell_Text()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("Click me");
        sheet["A1"].Hyperlink("https://example.com");
        sheet["A1"].GetString().Should().Be("Click me");
        sheet["A1"].GetHyperlink().Should().Be("https://example.com");
    }

    [Fact]
    public void Hyperlink_Returns_Same_Cell_For_Chaining()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        var cell = sheet["A1"];
        cell.Hyperlink("https://example.com").Should().BeSameAs(cell);
    }

    [Fact]
    public void GetHyperlink_Returns_Null_When_No_Hyperlink_Attached()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].GetHyperlink().Should().BeNull();
    }

    // ---- Verbatim target preservation ---------------------------------------

    [Fact]
    public void Hyperlink_Target_Is_Stored_Verbatim_Not_Canonicalized()
    {
        // The packaging layer writes the Uri's OriginalString into the .rels
        // Target — a mixed-case scheme/host must survive untouched (NPOI parity).
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                wb.AddSheet("S")["A1"].Hyperlink("HTTPS://Example.COM/MixedCase");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                wb["S"]["A1"].GetHyperlink().Should().Be("HTTPS://Example.COM/MixedCase");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Replace semantics ----------------------------------------------------

    [Fact]
    public void Hyperlink_Replaces_Existing_Hyperlink_On_Same_Cell()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Hyperlink("https://old.example.com");
        sheet["A1"].Hyperlink("https://new.example.com");

        sheet["A1"].GetHyperlink().Should().Be("https://new.example.com");

        // One <hyperlink> element and one relationship — the old pair is gone.
        var wsPart = SheetPart(wb);
        wsPart.Worksheet!.GetFirstChild<S.Hyperlinks>()!
            .Elements<S.Hyperlink>().Should().HaveCount(1);
        wsPart.HyperlinkRelationships.Should().HaveCount(1);
        wsPart.HyperlinkRelationships.Single().Uri.OriginalString
            .Should().Be("https://new.example.com");
    }

    [Fact]
    public void Replacing_External_With_Internal_Drops_The_Relationship()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Other");
        var sheet = wb.AddSheet("S");
        sheet["A1"].Hyperlink("https://example.com");
        sheet["A1"].Hyperlink("#Other!A1");

        sheet["A1"].GetHyperlink().Should().Be("Other!A1");
        SheetPart(wb).HyperlinkRelationships.Should().BeEmpty();
    }

    // ---- DOM shape -------------------------------------------------------------

    [Fact]
    public void External_Link_Uses_Relationship_And_Internal_Uses_Location()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Other");
        var sheet = wb.AddSheet("S");
        sheet["A1"].Hyperlink("https://example.com", display: "Example");
        sheet["A2"].Hyperlink("#Other!A1");

        var wsPart = wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts
            .Single(p => p.Worksheet!.Descendants<S.Hyperlink>().Any());
        var links = wsPart.Worksheet!.GetFirstChild<S.Hyperlinks>()!
            .Elements<S.Hyperlink>().ToList();

        var external = links.Single(h => h.Reference!.Value == "A1");
        external.Id!.Value.Should().NotBeNullOrEmpty();
        external.Location.Should().BeNull();
        external.Display!.Value.Should().Be("Example");
        wsPart.HyperlinkRelationships.Single(r => r.Id == external.Id!.Value)
            .IsExternal.Should().BeTrue();

        var internalLink = links.Single(h => h.Reference!.Value == "A2");
        internalLink.Id.Should().BeNull();
        internalLink.Location!.Value.Should().Be("Other!A1");
    }

    // ---- Round-trip -------------------------------------------------------------

    [Fact]
    public void Hyperlink_Roundtrips_Through_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var sheet = wb.AddSheet("S");
                sheet["A1"].Hyperlink("https://example.com", display: "Example");
                sheet["A2"].Hyperlink("mailto:foo@example.com");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                var sheet = wb["S"];
                sheet["A1"].GetHyperlink().Should().Be("https://example.com");
                sheet["A1"].GetString().Should().Be("Example");
                sheet["A2"].GetHyperlink().Should().Be("mailto:foo@example.com");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
