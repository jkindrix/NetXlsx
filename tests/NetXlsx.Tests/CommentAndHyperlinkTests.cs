// Coverage for the v0.7 sub-slice C cell-level annotation APIs:
// ICell.Comment / GetComment / GetCommentAuthor and
// ICell.Hyperlink / GetHyperlink.

using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class CommentAndHyperlinkTests
{
    // ---- Comments ----------------------------------------------------

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

    [Fact]
    public void Comment_Roundtrips_Through_Save_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"comment-rt-{Guid.NewGuid():N}.xlsx");
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

    // ---- Hyperlinks --------------------------------------------------

    [Theory]
    [InlineData("https://example.com")]
    [InlineData("http://example.com/path?q=1")]
    [InlineData("HTTPS://Example.COM")]                  // case-insensitive scheme
    [InlineData("mailto:user@example.com")]
    [InlineData("file:///c/temp/file.xlsx")]
    public void Hyperlink_Accepts_Supported_Schemes(string target)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Hyperlink(target);
        sheet["A1"].GetHyperlink().Should().Be(target);
    }

    [Fact]
    public void Hyperlink_Internal_Document_Form_Strips_Leading_Hash()
    {
        // NPOI stores the body without the '#' for Document-type links;
        // GetHyperlink returns the body verbatim.
        using var wb = Workbook.Create();
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
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].Hyperlink(target)))
            .Should().Throw<ArgumentException>()
            .WithMessage("*unsupported scheme*");
    }

    [Fact]
    public void Hyperlink_Null_Target_Throws_ArgumentNullException()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].Hyperlink(null!))).Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Hyperlink_Empty_Target_Throws_ArgumentException()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].Hyperlink(""))).Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Hyperlink_With_Display_Sets_Cell_Text()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Hyperlink("https://example.com", display: "Example");
        sheet["A1"].GetString().Should().Be("Example");
        sheet["A1"].GetHyperlink().Should().Be("https://example.com");
    }

    [Fact]
    public void Hyperlink_Without_Display_On_Empty_Cell_Sets_Text_To_Target()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].Hyperlink("https://example.com");
        sheet["A1"].GetString().Should().Be("https://example.com");
    }

    [Fact]
    public void Hyperlink_Without_Display_Preserves_Existing_Cell_Text()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("Click me");
        sheet["A1"].Hyperlink("https://example.com");
        sheet["A1"].GetString().Should().Be("Click me");
        sheet["A1"].GetHyperlink().Should().Be("https://example.com");
    }

    [Fact]
    public void Hyperlink_Returns_Same_Cell_For_Chaining()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var cell = sheet["A1"];
        cell.Hyperlink("https://example.com").Should().BeSameAs(cell);
    }

    [Fact]
    public void GetHyperlink_Returns_Null_When_No_Hyperlink_Attached()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].GetHyperlink().Should().BeNull();
    }

    [Fact]
    public void Hyperlink_Roundtrips_Through_Save_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"hyperlink-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("S");
                sheet["A1"].Hyperlink("https://example.com", display: "Example");
                sheet["A2"].Hyperlink("mailto:foo@example.com");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
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
