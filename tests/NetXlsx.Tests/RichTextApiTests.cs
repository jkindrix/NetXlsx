// Coverage for the v1.1 rich-text slice: value-type semantics for
// RichText / RichTextRun / RichTextStyle, plus end-to-end write +
// read-back through Workbook + Open.

using System;
using System.Collections.Generic;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class RichTextApiTests
{
    // ---- Value-type semantics -----------------------------------------

    [Fact]
    public void RichTextStyle_Default_Is_All_Null()
    {
        var s = RichTextStyle.Default;
        s.Bold.Should().BeNull();
        s.Italic.Should().BeNull();
        s.Underline.Should().BeNull();
        s.FontName.Should().BeNull();
        s.FontSize.Should().BeNull();
        s.Color.Should().BeNull();
    }

    [Fact]
    public void RichTextStyle_Equality_Is_Structural()
    {
        var a = new RichTextStyle { Bold = true, FontSize = 14, Color = Color.Red };
        var b = new RichTextStyle { Bold = true, FontSize = 14, Color = Color.Red };
        a.Should().Be(b);
        a.GetHashCode().Should().Be(b.GetHashCode());
    }

    [Fact]
    public void RichTextRun_Default_Style_Is_RichTextStyle_Default()
    {
        var r = new RichTextRun("hello");
        r.Text.Should().Be("hello");
        r.Style.Should().BeSameAs(RichTextStyle.Default);
    }

    [Fact]
    public void RichText_PlainText_Is_Concatenation_Of_Runs()
    {
        var rt = new RichText(
            new RichTextRun("Hello ", new RichTextStyle { Bold = true }),
            new RichTextRun("World", new RichTextStyle { Italic = true }));
        rt.PlainText.Should().Be("Hello World");
        rt.Runs.Should().HaveCount(2);
    }

    [Fact]
    public void RichText_Rejects_Empty_Runs_List()
    {
        Action act = () => _ = new RichText(Array.Empty<RichTextRun>());
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void RichText_Rejects_Null_Runs_List()
    {
        Action act = () => _ = new RichText((IReadOnlyList<RichTextRun>)null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void RichText_Rejects_Null_Run_In_List()
    {
        Action act = () => _ = new RichText(new RichTextRun[] { new("ok"), null! });
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void RichText_Equality_Is_Structural_Across_Runs()
    {
        var a = new RichText(
            new RichTextRun("a", new RichTextStyle { Bold = true }),
            new RichTextRun("b"));
        var b = new RichText(
            new RichTextRun("a", new RichTextStyle { Bold = true }),
            new RichTextRun("b"));
        a.Should().Be(b);
        a.GetHashCode().Should().Be(b.GetHashCode());
    }

    [Fact]
    public void RichText_Differs_When_Any_Run_Differs()
    {
        var a = new RichText(new RichTextRun("a", new RichTextStyle { Bold = true }));
        var b = new RichText(new RichTextRun("a", new RichTextStyle { Bold = false }));
        a.Should().NotBe(b);
    }

    // ---- Cell-level write + read --------------------------------------

    [Fact]
    public void SetRichText_Then_GetRichText_Roundtrips_In_Memory()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        var rt = new RichText(
            new RichTextRun("Bold ", new RichTextStyle { Bold = true }),
            new RichTextRun("italic ", new RichTextStyle { Italic = true }),
            new RichTextRun("red", new RichTextStyle { Color = Color.Red, Bold = true }));

        sh["A1"].SetRichText(rt);

        sh["A1"].Kind.Should().Be(CellKind.String);
        sh["A1"].GetString().Should().Be("Bold italic red");
        var read = sh["A1"].GetRichText();
        read.Should().NotBeNull();
        read!.PlainText.Should().Be("Bold italic red");
        // Run count matches what we wrote (no merging of adjacent equal styles).
        read.Runs.Should().HaveCount(3);
        read.Runs[0].Text.Should().Be("Bold ");
        read.Runs[0].Style.Bold.Should().Be(true);
        read.Runs[1].Style.Italic.Should().Be(true);
        read.Runs[2].Style.Color.Should().Be(Color.Red);
        read.Runs[2].Style.Bold.Should().Be(true);
    }

    [Fact]
    public void GetRichText_Returns_Null_For_Plain_String_Cell()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("plain");
        sh["A1"].GetRichText().Should().BeNull();
    }

    [Fact]
    public void GetRichText_Returns_Null_For_Non_String_Cells()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetNumber(42.0);
        sh["A2"].SetBool(true);
        sh["A3"].SetFormula("=1+1");
        sh["A1"].GetRichText().Should().BeNull();
        sh["A2"].GetRichText().Should().BeNull();
        sh["A3"].GetRichText().Should().BeNull();
    }

    [Fact]
    public void SetRichText_Rejects_Null_Argument()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh["A1"].SetRichText(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void SetRichText_Respects_MaxCellTextLength()
    {
        using var wb = Workbook.Create(new WorkbookOptions { MaxCellTextLength = 10 });
        var sh = wb.AddSheet("S");
        var rt = new RichText(
            new RichTextRun("0123456"),
            new RichTextRun("789X", new RichTextStyle { Bold = true }));  // total 11 > limit
        Action act = () => sh["A1"].SetRichText(rt);
        act.Should().Throw<ResourceLimitExceededException>();
    }

    [Fact]
    public void SetRichText_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"rt-roundtrip-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh["A1"].SetRichText(new RichText(
                    new RichTextRun("Hello ", new RichTextStyle { Bold = true, FontSize = 14 }),
                    new RichTextRun("World", new RichTextStyle { Italic = true, Color = Color.Blue })));
                wb.Save(path);
            }

            using (var wb = Workbook.Open(path))
            {
                var read = wb["S"]["A1"].GetRichText();
                read.Should().NotBeNull();
                read!.PlainText.Should().Be("Hello World");
                read.Runs.Should().HaveCount(2);
                read.Runs[0].Style.Bold.Should().Be(true);
                read.Runs[0].Style.FontSize.Should().Be(14);
                read.Runs[1].Style.Italic.Should().Be(true);
                read.Runs[1].Style.Color.Should().Be(Color.Blue);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void SetRichText_Skips_Zero_Length_Runs()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetRichText(new RichText(
            new RichTextRun("", new RichTextStyle { Bold = true }),
            new RichTextRun("real", new RichTextStyle { Italic = true })));
        var read = sh["A1"].GetRichText();
        read.Should().NotBeNull();
        read!.PlainText.Should().Be("real");
        // Empty leading run contributed no formatting; only the italic run remains.
        read.Runs.Should().HaveCount(1);
        read.Runs[0].Style.Italic.Should().Be(true);
    }

    // ---- Streaming-cell surface (type-honesty assertion) --------------

    [Fact]
    public void IStreamingCell_Does_Not_Expose_SetRichText()
    {
        // Decision I-50 / decision #7: NPOI's SXSSF SheetDataWriter
        // (NPOI 2.7.x) constructs a fresh XSSFRichTextString from the
        // plain string at flush time, dropping any in-memory formatting
        // runs. Rather than silently degrade or throw at runtime,
        // SetRichText is absent from IStreamingCell — type-honest.
        typeof(IStreamingCell)
            .GetMethod("SetRichText")
            .Should().BeNull(
                "rich-text writes through SXSSF lose formatting at the " +
                "NPOI serializer level; per decision #7 (streaming " +
                "type-honesty) the method is absent rather than silently " +
                "lossy or throw-on-call.");
    }
}
