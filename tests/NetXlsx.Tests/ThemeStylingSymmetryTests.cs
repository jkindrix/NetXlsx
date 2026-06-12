// I-89 — theme-color styling symmetry (XlsxCodeGen Appendix A #1): the
// font / rich-text-run / border-edge theme axes mirror the shipped
// CellStyle.BackgroundTheme exactly — write as <color theme tint>, theme
// wins over the literal color per slot, and the read path populates the
// theme property (literal null) for theme-indexed XML and vice versa.

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class ThemeStylingSymmetryTests
{
    private static readonly XNamespace M = SavedOoxml.Main;

    // An Excel-style high-precision tint — survives the OOXML double attr.
    private const double Tint = -0.249977111117893;

    // ---- Font color ----------------------------------------------------------

    [Fact]
    public void FontColorTheme_Writes_Theme_And_Tint_Attributes()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].Style(new CellStyle { FontColorTheme = new ThemeColor(4, Tint) });

        var styles = SavedOoxml.StylesXml(wb);
        var color = styles.Root!.Element(M + "fonts")!.Elements(M + "font").Last()
            .Element(M + "color")!;
        ((string?)color.Attribute("theme")).Should().Be("4");
        ((double?)color.Attribute("tint")).Should().BeApproximately(Tint, 1e-15);
        ((string?)color.Attribute("rgb")).Should().BeNull();
    }

    [Fact]
    public void FontColorTheme_Zero_Tint_Omits_The_Tint_Attribute()
    {
        // Same convention the I-79 fill path established.
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].Style(new CellStyle { FontColorTheme = new ThemeColor(4) });
        var color = SavedOoxml.StylesXml(wb).Root!
            .Element(M + "fonts")!.Elements(M + "font").Last().Element(M + "color")!;
        ((string?)color.Attribute("theme")).Should().Be("4");
        color.Attribute("tint").Should().BeNull();
    }

    [Fact]
    public void FontColorTheme_Wins_Over_FontColor_When_Both_Set()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].Style(new CellStyle
        {
            FontColorTheme = new ThemeColor(4),
            FontColor = Color.FromHex("#FF0000"),
        });
        var color = SavedOoxml.StylesXml(wb).Root!
            .Element(M + "fonts")!.Elements(M + "font").Last().Element(M + "color")!;
        ((string?)color.Attribute("theme")).Should().Be("4");
        ((string?)color.Attribute("rgb")).Should().BeNull("the theme variant wins over the literal (I-79 rule)");
    }

    [Fact]
    public void FontColorTheme_Round_Trips_Through_A_File()
    {
        var path = Path.Combine(Path.GetTempPath(), $"i89-font-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S")["A1"].Style(new CellStyle { FontColorTheme = new ThemeColor(4, Tint) });
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var style = wb["S"]["A1"].GetStyle();
                style.FontColorTheme.Should().Be(new ThemeColor(4, Tint));
                style.FontColor.Should().BeNull("a theme-indexed color reads via the theme axis only");
            }
        }
        finally { File.Delete(path); }
    }

    // ---- Rich-text run color --------------------------------------------------

    [Fact]
    public void Run_ColorTheme_Round_Trips_And_Wins_Over_Literal()
    {
        var path = Path.Combine(Path.GetTempPath(), $"i89-run-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S")["A1"].SetRichText(new RichText(
                    new RichTextRun("hot", new RichTextStyle
                    {
                        ColorTheme = new ThemeColor(5, Tint),
                        Color = Color.FromHex("#FF0000"), // loses to the theme
                    }),
                    new RichTextRun("cold", new RichTextStyle { Color = Color.FromHex("#0000FF") })));
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var rt = wb["S"]["A1"].GetRichText()!;
                rt.Runs[0].Style.ColorTheme.Should().Be(new ThemeColor(5, Tint));
                rt.Runs[0].Style.Color.Should().BeNull("theme-indexed run colors read via ColorTheme only");
                rt.Runs[1].Style.ColorTheme.Should().BeNull();
                rt.Runs[1].Style.Color.Should().Be(Color.FromHex("#0000FF"));
            }
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Run_ColorTheme_Writes_Theme_Attribute_In_The_Cell_Run()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetRichText(new RichText(
            new RichTextRun("x", new RichTextStyle { ColorTheme = new ThemeColor(5) })));
        var sheet = SavedOoxml.SheetXml(wb);
        var rpr = SavedOoxml.Cell(sheet, "A1")!
            .Element(M + "is")!.Elements(M + "r").First().Element(M + "rPr")!;
        ((string?)rpr.Element(M + "color")!.Attribute("theme")).Should().Be("5");
    }

    // ---- Border edges -----------------------------------------------------------

    [Fact]
    public void Border_Edge_Themes_Round_Trip_Per_Edge()
    {
        var path = Path.Combine(Path.GetTempPath(), $"i89-border-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S")["A1"].Style(new CellStyle
                {
                    Borders = new CellBorders(
                        Top: BorderStyle.Thin,
                        Bottom: BorderStyle.Double, BottomColor: Color.FromHex("#00FF00"),
                        Left: BorderStyle.Dashed, LeftColor: Color.FromHex("#FF0000"))
                    {
                        TopColorTheme = new ThemeColor(4, Tint),
                        // Left carries BOTH: the theme must win on that edge only.
                        LeftColorTheme = new ThemeColor(6),
                    },
                });
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var b = wb["S"]["A1"].GetStyle().Borders!;
                b.Top.Should().Be(BorderStyle.Thin);
                b.TopColorTheme.Should().Be(new ThemeColor(4, Tint));
                b.TopColor.Should().BeNull();
                b.Bottom.Should().Be(BorderStyle.Double);
                b.BottomColorTheme.Should().BeNull("a literal edge stays literal");
                b.BottomColor.Should().Be(Color.FromHex("#00FF00"));
                b.Left.Should().Be(BorderStyle.Dashed);
                b.LeftColorTheme.Should().Be(new ThemeColor(6));
                b.LeftColor.Should().BeNull("the theme variant wins per edge");
                b.Right.Should().BeNull();
            }
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Border_Edge_Theme_Writes_Theme_Attribute_In_Styles_Xml()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].Style(new CellStyle
        {
            Borders = new CellBorders(Top: BorderStyle.Thin) { TopColorTheme = new ThemeColor(4) },
        });
        var border = SavedOoxml.StylesXml(wb).Root!
            .Element(M + "borders")!.Elements(M + "border").Last();
        var topColor = border.Element(M + "top")!.Element(M + "color")!;
        ((string?)topColor.Attribute("theme")).Should().Be("4");
        ((string?)topColor.Attribute("rgb")).Should().BeNull();
    }

    // ---- Pool behavior ------------------------------------------------------------

    [Fact]
    public void Equal_Theme_Styles_Share_One_Xf_And_Distinct_Tints_Do_Not()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].Style(new CellStyle { FontColorTheme = new ThemeColor(4, 0.4) });
        s["A2"].Style(new CellStyle { FontColorTheme = new ThemeColor(4, 0.4) });
        s["A3"].Style(new CellStyle { FontColorTheme = new ThemeColor(4, 0.6) });

        var sheet = SavedOoxml.SheetXml(wb);
        var a1 = SavedOoxml.CellStyleIndex(sheet, "A1");
        var a2 = SavedOoxml.CellStyleIndex(sheet, "A2");
        var a3 = SavedOoxml.CellStyleIndex(sheet, "A3");
        a1.Should().NotBeNull();
        a2.Should().Be(a1, "structurally equal theme styles dedup to one xf (decision #4)");
        a3.Should().NotBe(a1, "a different tint is a different font");
    }

    [Fact]
    public void Named_Style_With_Theme_Axes_Rehydrates_After_Reopen()
    {
        var path = Path.Combine(Path.GetTempPath(), $"i89-named-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            var brand = new CellStyle
            {
                FontColorTheme = new ThemeColor(4),
                Borders = new CellBorders(Top: BorderStyle.Thin) { TopColorTheme = new ThemeColor(5, 0.2) },
            };
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S");
                wb.RegisterStyle("Brand", brand);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var back = wb.GetRegisteredStyle("Brand")!;
                back.FontColorTheme.Should().Be(new ThemeColor(4));
                back.Borders!.TopColorTheme.Should().Be(new ThemeColor(5, 0.2));
            }
        }
        finally { File.Delete(path); }
    }

    // ---- Truthful read-back of the scaffolding font --------------------------------

    [Fact]
    public void Style_Read_From_A_Default_Font_Xf_Reports_Theme1_And_Reapplying_It_Embeds()
    {
        // A style with no font axes references font 0 — whose color element
        // genuinely IS <color theme="1"/> (the Excel-conventional default
        // text color). The read path reports theme attributes truthfully
        // (suppressing index 1 would misreport legitimate dk1-tinted fonts),
        // so the read-back carries FontColorTheme = (1, 0) alongside the
        // already-shipped font-0 name/size leak — and REAPPLYING that read
        // style is a theme-indexed write, which embeds the default theme.
        // Deliberate, pinned consequence of truthful read-back (I-89): the
        // copied output renders identically; it is no longer theme-less.
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].Style(new CellStyle { Background = Color.FromHex("#FFEEDD") });

        var read = s["A1"].GetStyle();
        read.FontColorTheme.Should().Be(new ThemeColor(1));
        read.FontColor.Should().BeNull();
        wb.GetThemeXml().Should().BeNull("reading styles must not embed anything");

        s["A2"].Style(read);
        wb.GetThemeXml().Should().NotBeNull("re-applying a theme-carrying read style is a theme write");
    }

    // ---- Foreign-authored theme+tint read-back ------------------------------------
    // Excel writes e.g. <color theme="5" tint="-0.249977111117893"/> on fonts;
    // simulate a third-party author by post-editing a saved file's styles.xml
    // (the writer below is NOT NetXlsx's emission path).

    [Fact]
    public void Foreign_Theme_Tint_Font_Color_Reads_Back()
    {
        var path = Path.Combine(Path.GetTempPath(), $"i89-foreign-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                // A literal-color font allocates font index 1; the post-edit
                // below swaps its color element for a theme+tint one.
                wb.AddSheet("S")["A1"].Style(new CellStyle { FontColor = Color.FromHex("#123456") });
                wb.Save(path);
            }

            using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                var entry = zip.GetEntry("xl/styles.xml")!;
                string xml;
                using (var r = new StreamReader(entry.Open())) xml = r.ReadToEnd();
                // Swap the attribute only — robust to the SDK's element
                // prefix/spacing choices (it writes <x:color rgb="..." />).
                xml.Should().Contain("rgb=\"FF123456\"", "the literal font color must be present to post-edit");
                xml = xml.Replace("rgb=\"FF123456\"",
                    "theme=\"5\" tint=\"-0.249977111117893\"");
                entry.Delete();
                var fresh = zip.CreateEntry("xl/styles.xml");
                using var w = new StreamWriter(fresh.Open());
                w.Write(xml);
            }

            using (var wb = Workbook.Open(path))
            {
                var style = wb["S"]["A1"].GetStyle();
                style.FontColorTheme.Should().Be(new ThemeColor(5, -0.249977111117893));
                style.FontColor.Should().BeNull();
            }
        }
        finally { File.Delete(path); }
    }
}
