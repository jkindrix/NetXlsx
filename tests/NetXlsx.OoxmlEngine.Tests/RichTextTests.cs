// I-82 engine swap — rich-text slice conformance.
//
// Covers ICell.SetRichText / GetRichText on the Open XML SDK engine: multi-run
// inline rich strings, the marquee inheritance semantic (an empty-style run gets
// NO <rPr> and so inherits the cell font — lesson #10), full font-axis round-trip,
// Kind/GetString behavior for rich-text cells, empty-run handling, the text-length
// limit, and the "plain string / non-string cell -> null" GetRichText contract.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using S = DocumentFormat.OpenXml.Spreadsheet;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class RichTextTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-richtext-{Guid.NewGuid():N}.xlsx");

    private static RichText Sample() => new(
        new RichTextRun("VERY IMPORTANT", RichTextStyle.Default),
        new RichTextRun(" please read", new RichTextStyle { Bold = true, Color = Color.FromRgb(0xFF, 0, 0) }));

    // ---- Round-trip ---------------------------------------------------------

    [Fact]
    public void SetRichText_Round_Trips_Multiple_Runs_Through_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                wb.AddSheet("S")["A1"].SetRichText(Sample());
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                wb["S"]["A1"].GetRichText().Should().Be(Sample());
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Full_Font_Axes_Round_Trip()
    {
        var path = TempXlsxPath();
        try
        {
            var style = new RichTextStyle
            {
                Bold = true,
                Italic = true,
                Underline = UnderlineStyle.Single,
                FontName = "Arial",
                FontSize = 14,
                Color = Color.FromRgb(0x10, 0x20, 0x30),
            };
            using (var wb = Workbook.CreateOoxml())
            {
                wb.AddSheet("S")["A1"].SetRichText(new RichText(new RichTextRun("x", style)));
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                var run = wb["S"]["A1"].GetRichText()!.Runs.Single();
                run.Text.Should().Be("x");
                run.Style.Should().Be(style);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Marquee: empty-style run inherits (no <rPr>) -----------------------

    [Fact]
    public void Empty_Style_Run_Writes_No_RunProperties_So_It_Inherits()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S")["A1"].SetRichText(Sample());

        var ws = wb.Underlying.WorkbookPart!.WorksheetParts.First().Worksheet!;
        var cell = ws.Descendants<S.Cell>().First(c => c.CellReference == "A1");
        var runs = cell.InlineString!.Elements<S.Run>().ToList();

        runs.Should().HaveCount(2);
        runs[0].RunProperties.Should().BeNull("an empty-style run carries no <rPr> and inherits the cell font (lesson #10)");
        runs[1].RunProperties.Should().NotBeNull("a formatted run carries its <rPr>");
    }

    [Fact]
    public void Empty_Style_Run_Reads_Back_As_Default_Style()
    {
        using var wb = Workbook.CreateOoxml();
        var c = wb.AddSheet("S")["A1"];
        c.SetRichText(Sample());

        var rt = c.GetRichText()!;
        rt.Runs[0].Style.Should().Be(RichTextStyle.Default);
        rt.Runs[1].Style.Bold.Should().BeTrue();
        rt.Runs[1].Style.Color.Should().Be(Color.FromRgb(0xFF, 0, 0));
    }

    // ---- Kind / GetString ---------------------------------------------------

    [Fact]
    public void RichText_Cell_Is_A_String_Kind_With_Concatenated_Text()
    {
        using var wb = Workbook.CreateOoxml();
        var c = wb.AddSheet("S")["A1"];
        c.SetRichText(Sample());

        c.Kind.Should().Be(CellKind.String);
        c.GetString().Should().Be("VERY IMPORTANT please read");
    }

    // ---- GetRichText null contract -----------------------------------------

    [Fact]
    public void Plain_String_Cell_Returns_Null_RichText()
    {
        using var wb = Workbook.CreateOoxml();
        var c = wb.AddSheet("S")["A1"];
        c.SetString("just a string");
        c.GetRichText().Should().BeNull();
    }

    [Fact]
    public void Non_String_Cell_Returns_Null_RichText()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetNumber(42);
        s["A2"].SetBool(true);
        s["A3"].SetDate(new DateTime(2026, 5, 31));
        s["A1"].GetRichText().Should().BeNull();
        s["A2"].GetRichText().Should().BeNull();
        s["A3"].GetRichText().Should().BeNull();
    }

    [Fact]
    public void Empty_Cell_Returns_Null_RichText()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S")["Z99"].GetRichText().Should().BeNull();
    }

    // ---- Edge cases ---------------------------------------------------------

    [Fact]
    public void Empty_Runs_Contribute_No_Formatting_Run()
    {
        using var wb = Workbook.CreateOoxml();
        var c = wb.AddSheet("S")["A1"];
        // The middle run has empty text and a (would-be) style; it must be skipped.
        c.SetRichText(new RichText(
            new RichTextRun("a", new RichTextStyle { Bold = true }),
            new RichTextRun("", new RichTextStyle { Italic = true }),
            new RichTextRun("b", new RichTextStyle { Italic = true })));

        var rt = c.GetRichText()!;
        rt.Runs.Should().HaveCount(2);
        rt.PlainText.Should().Be("ab");
        rt.Runs[0].Style.Bold.Should().BeTrue();
        rt.Runs[1].Style.Italic.Should().BeTrue();
    }

    [Fact]
    public void Whitespace_In_Runs_Is_Preserved()
    {
        using var wb = Workbook.CreateOoxml();
        var c = wb.AddSheet("S")["A1"];
        c.SetRichText(new RichText(
            new RichTextRun("  lead", new RichTextStyle { Bold = true }),
            new RichTextRun("trail  ", RichTextStyle.Default)));
        c.GetString().Should().Be("  leadtrail  ");
    }

    [Fact]
    public void SetRichText_Overwrites_A_Prior_Value()
    {
        using var wb = Workbook.CreateOoxml();
        var c = wb.AddSheet("S")["A1"];
        c.SetNumber(123);
        c.SetRichText(Sample());
        c.Kind.Should().Be(CellKind.String);
        c.GetNumber().Should().BeNull();
        c.GetRichText().Should().Be(Sample());
    }

    [Fact]
    public void SetRichText_Null_Throws()
    {
        using var wb = Workbook.CreateOoxml();
        var c = wb.AddSheet("S")["A1"];
        ((Action)(() => c.SetRichText(null!))).Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void SetRichText_Exceeding_Text_Limit_Throws()
    {
        using var wb = Workbook.CreateOoxml(new WorkbookOptions { MaxCellTextLength = 10 });
        var c = wb.AddSheet("S")["A1"];
        var rt = new RichText(new RichTextRun(new string('x', 11)));
        ((Action)(() => c.SetRichText(rt))).Should().Throw<ResourceLimitExceededException>();
    }
}
