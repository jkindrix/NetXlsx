// I-89 — pins the facts of the embedded default Office theme
// (Workbook.DefaultThemeXml / Internal/DefaultTheme.cs) against the values
// transcribed from the Excel-authored provenance source: the clrScheme slot
// colors, the font scheme, the lnStyleLst widths, and the two non-ASCII
// typeface names (carried as \uXXXX escapes in the constant so a source
// transcoding accident fails HERE instead of corrupting silently — the
// expected strings below are escape-spelled too, so the two files cannot
// drift together).

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class DefaultThemeFactsTests
{
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static XElement ThemeElements()
    {
        var doc = XDocument.Load(new MemoryStream(Workbook.DefaultThemeXml));
        doc.Root!.Name.Should().Be(A + "theme");
        doc.Root!.Attribute("name")!.Value.Should().Be("Office Theme");
        return doc.Root!.Element(A + "themeElements")!;
    }

    [Theory]
    // The current Office color scheme — dk1/lt1 are system colors with
    // explicit lastClr fallbacks; everything else is literal sRGB.
    [InlineData("dk2", "44546A")]
    [InlineData("lt2", "E7E6E6")]
    [InlineData("accent1", "4472C4")] // "Office blue" — the I-89 acceptance value
    [InlineData("accent2", "ED7D31")]
    [InlineData("accent3", "A5A5A5")]
    [InlineData("accent4", "FFC000")]
    [InlineData("accent5", "5B9BD5")]
    [InlineData("accent6", "70AD47")]
    [InlineData("hlink", "0563C1")]
    [InlineData("folHlink", "954F72")]
    public void ClrScheme_Srgb_Slot_Has_The_Office_Value(string slot, string expectedHex)
    {
        var clrScheme = ThemeElements().Element(A + "clrScheme")!;
        clrScheme.Attribute("name")!.Value.Should().Be("Office");
        clrScheme.Element(A + slot)!.Element(A + "srgbClr")!
            .Attribute("val")!.Value.Should().Be(expectedHex);
    }

    [Theory]
    [InlineData("dk1", "windowText", "000000")]
    [InlineData("lt1", "window", "FFFFFF")]
    public void ClrScheme_System_Slot_Has_The_Office_Value(string slot, string sysVal, string lastClr)
    {
        var el = ThemeElements().Element(A + "clrScheme")!.Element(A + slot)!.Element(A + "sysClr")!;
        el.Attribute("val")!.Value.Should().Be(sysVal);
        el.Attribute("lastClr")!.Value.Should().Be(lastClr);
    }

    [Fact]
    public void FontScheme_Is_Calibri_Light_Over_Calibri()
    {
        var fontScheme = ThemeElements().Element(A + "fontScheme")!;
        fontScheme.Element(A + "majorFont")!.Element(A + "latin")!
            .Attribute("typeface")!.Value.Should().Be("Calibri Light");
        // The minor font MUST stay Calibri: created-workbook stylesheets mark
        // font 0 <scheme val="minor"/>, so consumers resolve the default cell
        // font through this slot once the theme embeds (DefaultTheme.cs header).
        fontScheme.Element(A + "minorFont")!.Element(A + "latin")!
            .Attribute("typeface")!.Value.Should().Be("Calibri");
    }

    [Theory]
    [InlineData("majorFont")]
    [InlineData("minorFont")]
    public void NonAscii_Script_Typefaces_Survive_Encoding(string collection)
    {
        // Escape-spelled expectations (Hang = Malgun Gothic, Hant = PMingLiU
        // in their native spellings) — see the file header for why.
        var fonts = ThemeElements().Element(A + "fontScheme")!.Element(A + collection)!;
        string TypefaceOf(string script) => fonts.Elements(A + "font")
            .First(f => f.Attribute("script")?.Value == script)
            .Attribute("typeface")!.Value;
        TypefaceOf("Hang").Should().Be("맑은 고딕");
        TypefaceOf("Hant").Should().Be("新細明體");
    }

    [Fact]
    public void LnStyleLst_Has_The_Office_Widths()
    {
        var widths = ThemeElements().Element(A + "fmtScheme")!.Element(A + "lnStyleLst")!
            .Elements(A + "ln").Select(ln => (string?)ln.Attribute("w")).ToArray();
        widths.Should().Equal("6350", "12700", "19050");
    }

    [Fact]
    public void Resolves_Through_The_Engine_Like_Any_Theme()
    {
        using var wb = Workbook.Create();
        wb.SetThemeXml(Workbook.DefaultThemeXml);
        wb.ResolveThemeColor(4)!.Value.ToHex().Should().Be("#FF4472C4"); // accent1
        wb.ResolveThemeColor("bg2")!.Value.ToHex().Should().Be("#FFE7E6E6"); // lt2 alias
        wb.GetThemeLineWidthEmu(1).Should().Be(6350);
        wb.GetThemeLineWidthEmu(3).Should().Be(19050);
    }

    [Fact]
    public void Each_Call_Returns_An_Isolated_Fresh_Copy()
    {
        var a = Workbook.DefaultThemeXml;
        var b = Workbook.DefaultThemeXml;
        ReferenceEquals(a, b).Should().BeFalse("each call must return a new array");
        a[0] = 0xFF;
        Workbook.DefaultThemeXml[0].Should().Be((byte)'<', "mutating a returned copy must not leak");
    }
}
