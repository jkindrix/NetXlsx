// Coverage for the v0.4 styling slice: Color, CellStyle, NumberFormats,
// ICell.Style/NumberFormat/GetStyle, and the style-pool dedup behavior.

using System;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class StyleApiTests
{
    // ---- Color value semantics -----------------------------------------

    [Fact]
    public void Color_ARGB_Equality_Is_Structural()
    {
        // Decision I-23: equal ARGB bytes -> equal Color, regardless of construction.
        Color.FromRgb(255, 0, 0).Should().Be(Color.FromArgb(0xFF, 255, 0, 0));
        Color.Red.Should().Be(Color.FromRgb(255, 0, 0));
    }

    [Theory]
    [InlineData("#FF0000",   0xFF, 0xFF, 0x00, 0x00)]
    [InlineData("#80FF0000", 0x80, 0xFF, 0x00, 0x00)]
    [InlineData("FF0000",    0xFF, 0xFF, 0x00, 0x00)]   // # is optional
    public void Color_FromHex_Parses_Both_Forms(string hex, byte a, byte r, byte g, byte b)
    {
        var c = Color.FromHex(hex);
        c.A.Should().Be(a);
        c.R.Should().Be(r);
        c.G.Should().Be(g);
        c.B.Should().Be(b);
    }

    [Theory]
    [InlineData("")]
    [InlineData("ABC")]
    [InlineData("#GGGGGG")]
    public void Color_FromHex_Rejects_Bad_Input(string hex)
    {
        Action act = () => Color.FromHex(hex);
        act.Should().Throw<Exception>();
    }

    [Fact]
    public void Color_ToHex_Roundtrips_Through_FromHex()
    {
        var c = Color.FromArgb(0x12, 0x34, 0x56, 0x78);
        var hex = c.ToHex();
        hex.Should().Be("#12345678");
        Color.FromHex(hex).Should().Be(c);
    }

    // ---- CellStyle value semantics -------------------------------------

    [Fact]
    public void CellStyle_Default_Is_All_Null()
    {
        var s = CellStyle.Default;
        s.Bold.Should().BeNull();
        s.Background.Should().BeNull();
        s.NumberFormat.Should().BeNull();
        // … all-null by design.
    }

    [Fact]
    public void CellStyle_Equality_Is_Structural()
    {
        var a = new CellStyle { Bold = true, NumberFormat = "0.00" };
        var b = new CellStyle { Bold = true, NumberFormat = "0.00" };
        var c = new CellStyle { Bold = true, NumberFormat = "0.00", Italic = true };
        a.Should().Be(b);
        a.Should().NotBe(c);
    }

    // ---- Apply / round-trip via the pool -------------------------------

    [Fact]
    public void Style_Applies_NumberFormat_To_Cell()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetNumber(1234.56);
        sheet["A1"].Style(new CellStyle { NumberFormat = NumberFormats.Currency });

        sheet["A1"].GetStyle().NumberFormat.Should().Be(NumberFormats.Currency);
    }

    [Fact]
    public void NumberFormat_Shortcut_Is_Equivalent_To_Style_With_NumberFormat()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetNumber(1234.56);
        sheet["A1"].NumberFormat(NumberFormats.NumberTwo);

        sheet["A1"].GetStyle().NumberFormat.Should().Be(NumberFormats.NumberTwo);
    }

    [Fact]
    public void Style_Bold_Applies_Font_And_Survives_Read()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("Header");
        sheet["A1"].Style(new CellStyle { Bold = true });

        // The persisted stylesheet carries a bold font…
        SavedOoxml.StylesXml(wb).Descendants(SavedOoxml.Main + "font")
            .Should().Contain(f => f.Element(SavedOoxml.Main + "b") != null,
                "a bold cell font must persist as <font><b/>");

        // …and the public read-back agrees.
        var roundtripped = sheet["A1"].GetStyle();
        roundtripped.Bold.Should().Be(true);
    }

    [Fact]
    public void Style_Background_Color_Applies_Solid_Fill()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("hi");
        sheet["A1"].Style(new CellStyle { Background = Color.LightGray });

        // The persisted stylesheet carries a solid pattern fill…
        SavedOoxml.StylesXml(wb).Descendants(SavedOoxml.Main + "patternFill")
            .Should().Contain(p => (string?)p.Attribute("patternType") == "solid",
                "a background color must persist as a solid patternFill");

        // …and the public read-back agrees.
        var rt = sheet["A1"].GetStyle();
        rt.Background.Should().Be(Color.LightGray);
    }

    [Fact]
    public void Style_Merges_Over_Existing_Style()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetNumber(99.50m);
        sheet["A1"].Style(new CellStyle { NumberFormat = NumberFormats.Currency });
        // Second call adds Bold; should NOT clobber NumberFormat.
        sheet["A1"].Style(new CellStyle { Bold = true });

        var rt = sheet["A1"].GetStyle();
        rt.NumberFormat.Should().Be(NumberFormats.Currency);
        rt.Bold.Should().Be(true);
    }

    [Fact]
    public void Style_Pool_Dedupes_Equal_Styles_Across_Many_Cells()
    {
        // 100 cells styled identically should share ONE ICellStyle instance.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var style = new CellStyle { Bold = true, NumberFormat = NumberFormats.Currency };

        for (int i = 1; i <= 100; i++)
        {
            sheet[i, 1].SetNumber(i);
            sheet[i, 1].Style(style);
        }

        var sheetXml = SavedOoxml.SheetXml(wb);
        var firstIdx = SavedOoxml.CellStyleIndex(sheetXml, "A1");
        firstIdx.Should().NotBeNull("the styled cell must carry an explicit style index");
        for (int i = 2; i <= 100; i++)
        {
            SavedOoxml.CellStyleIndex(sheetXml, $"A{i}").Should().Be(firstIdx,
                "structurally-equal styles must share one cellXfs entry (decision #4)");
        }
    }

    [Fact]
    public void DateTime_Default_Style_Flows_Through_The_Same_Pool()
    {
        // S29 absorption: the date/time default styles are now pool
        // entries. Two cells set via SetDate with no prior style should
        // share the same ICellStyle index.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetDate(new DateOnly(2026, 5, 16));
        sheet["A2"].SetDate(new DateOnly(2026, 5, 17));

        var sheetXml = SavedOoxml.SheetXml(wb);
        var a1 = SavedOoxml.CellStyleIndex(sheetXml, "A1");
        a1.Should().NotBeNull("the date default style must persist on the cell");
        SavedOoxml.CellStyleIndex(sheetXml, "A2").Should().Be(a1,
            "the date default style is allocated through the pool and dedup-shared");
    }

    [Fact]
    public void GetStyle_On_Unstyled_Cell_Returns_Default()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("plain");
        sheet["A1"].GetStyle().Should().Be(CellStyle.Default);
    }

    [Fact]
    public void Style_Throws_On_Null()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        Action act = () => sheet["A1"].Style(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Style_Returns_Self_For_Chaining()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var cell = sheet["A1"];
        cell.Style(new CellStyle { Bold = true }).Should().BeSameAs(cell);
        cell.NumberFormat("0.00").Should().BeSameAs(cell);
    }
}
