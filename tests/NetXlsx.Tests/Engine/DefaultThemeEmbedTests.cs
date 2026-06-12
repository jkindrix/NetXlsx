// I-89 — the lazy default-theme embed. A theme-indexed color written into a
// workbook that has no theme part embeds Workbook.DefaultThemeXml first (the
// EnsureThemePart choke point / the streaming engine's assembly-time check),
// so theme references resolve consumer-independently instead of against
// whatever theme each consumer substitutes (the R-8 lottery, proven with LO).
// Workbooks that never write theme-indexed styling must stay theme-free —
// and therefore byte-identical to pre-I-89 output.

using System;
using System.IO;
using System.IO.Compression;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class DefaultThemeEmbedTests
{
    // 1x1 transparent PNG (same fixture as ThemeReadAndDrawingIterationTests).
    private static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

    /// <summary>True when the saved package contains xl/theme/theme1.xml.</summary>
    private static bool SavedThemePartExists(IWorkbook wb)
    {
        using var ms = new MemoryStream();
        wb.Save(ms);
        ms.Position = 0;
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        return zip.GetEntry("xl/theme/theme1.xml") is not null;
    }

    private static void AssertEmbedded(IWorkbook wb)
    {
        wb.GetThemeXml().Should().NotBeNull("a theme-indexed write must embed the default theme");
        // Office accent1 — the documented I-81 contract starts holding on fresh output.
        wb.ResolveThemeColor(4)!.Value.ToHex().Should().Be("#FF4472C4");
        SavedThemePartExists(wb).Should().BeTrue("the embedded theme must persist as xl/theme/theme1.xml");
    }

    // ---- Plain workbooks stay theme-free --------------------------------

    [Fact]
    public void Plain_Workbook_With_Literal_Styles_Gets_No_Theme_Part()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("plain");
        // Every non-theme style family: font axes (allocates a font element —
        // proving the scaffolding font-0 <color theme="1"/> exemption holds),
        // RGB fill, RGB-colored borders, number format, alignment.
        s["A2"].Style(new CellStyle
        {
            Bold = true,
            FontColor = Color.FromHex("#336699"),
            Background = Color.FromHex("#EEEEEE"),
            Borders = CellBorders.All(BorderStyle.Thin, Color.FromRgb(1, 2, 3)),
            NumberFormat = "0.00",
            HorizontalAlignment = HAlign.Center,
        });
        wb.ResolveThemeColor(4).Should().BeNull("the I-81 no-theme contract is unchanged");
        wb.GetThemeXml().Should().BeNull();
        SavedThemePartExists(wb).Should().BeFalse("theme-free workbooks must stay byte-identical");
    }

    // ---- Style-pool trigger sites ----------------------------------------

    [Fact]
    public void FontColorTheme_Embeds_The_Default_Theme()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].Style(new CellStyle { FontColorTheme = new ThemeColor(4) });
        AssertEmbedded(wb);
    }

    [Fact]
    public void BackgroundTheme_Embeds_The_Default_Theme()
    {
        // The shipped I-79 axis participates too — its outputs stop being a
        // consumer lottery the same way.
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].Style(new CellStyle { BackgroundTheme = new ThemeColor(5, 0.4) });
        AssertEmbedded(wb);
    }

    [Fact]
    public void Border_Edge_Theme_Embeds_The_Default_Theme()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].Style(new CellStyle
        {
            Borders = new CellBorders(Top: BorderStyle.Thin) { TopColorTheme = new ThemeColor(6) },
        });
        AssertEmbedded(wb);
    }

    [Fact]
    public void Border_Theme_On_A_StyleLess_Edge_Is_Inert()
    {
        // A color (literal or theme) on an edge with no BorderStyle emits
        // nothing — same contract as the literal edge colors — so it must
        // not drag a theme part in.
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].Style(new CellStyle
        {
            Borders = new CellBorders(Top: BorderStyle.Thin, TopColor: Color.Black)
            {
                BottomColorTheme = new ThemeColor(4), // no Bottom style
            },
        });
        wb.GetThemeXml().Should().BeNull();
    }

    [Fact]
    public void RichText_Run_ColorTheme_Embeds_The_Default_Theme()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetRichText(new RichText(
            new RichTextRun("themed", new RichTextStyle { ColorTheme = new ThemeColor(4) })));
        AssertEmbedded(wb);
    }

    [Fact]
    public void RichText_Run_ColorTheme_On_An_Empty_Run_Is_Inert()
    {
        // Empty runs are skipped at emission and contribute no <r>, so their
        // style must not drag a theme part in.
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetRichText(new RichText(
            new RichTextRun("plain"),
            new RichTextRun("", new RichTextStyle { ColorTheme = new ThemeColor(4) })));
        wb.GetThemeXml().Should().BeNull();
    }

    // ---- Drawing-layer trigger sites --------------------------------------

    [Fact]
    public void Theme_Picture_Border_Embeds_The_Default_Theme()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var pic = s.AddPicture("B2", "D6", OnePixelPng, ImageFormat.Png);
        pic.Border = new PictureBorder { ThemeColor = new ThemeColor(5), WidthPoints = 1 };
        AssertEmbedded(wb);
    }

    [Fact]
    public void Rgb_Picture_Border_Is_Theme_Free()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var pic = s.AddPicture("B2", "D6", OnePixelPng, ImageFormat.Png);
        pic.Border = new PictureBorder { Color = Color.FromHex("#FF0000"), WidthPoints = 1 };
        wb.GetThemeXml().Should().BeNull();
    }

    [Fact]
    public void Connector_Embeds_The_Default_Theme()
    {
        // AddConnector always writes the NPOI-parity <xdr:style> block —
        // lnRef/fillRef/effectRef accent1 + fontRef tx1 are theme-indexed
        // writes whether the user picked them or not.
        using var wb = Workbook.Create();
        wb.AddSheet("S").AddConnector(ConnectorType.Straight, "A1", "C3");
        AssertEmbedded(wb);
    }

    [Fact]
    public void Pie_Chart_Embeds_The_Default_Theme()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("a"); s["A2"].SetString("b");
        s["B1"].SetNumber(1); s["B2"].SetNumber(2);
        s.AddChart(ChartType.Pie, "D1", "K12", "A1:A2", "B1:B2");
        AssertEmbedded(wb);
    }

    [Fact]
    public void Column_Chart_Is_Theme_Free()
    {
        // Only the pie path emits scheme colors (accent-cycled slice fills);
        // the other chart types write none and must not trigger.
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("a"); s["A2"].SetString("b");
        s["B1"].SetNumber(1); s["B2"].SetNumber(2);
        s.AddChart(ChartType.Column, "D1", "K12", "A1:A2", "B1:B2");
        wb.GetThemeXml().Should().BeNull();
    }

    [Fact]
    public void Plain_Shape_Is_Theme_Free()
    {
        // AddShape emits literal RGB / noFill only — no scheme colors.
        using var wb = Workbook.Create();
        wb.AddSheet("S").AddShape(ShapeType.Rectangle, "A1", "C3",
            fillColor: Color.FromHex("#CCCCCC"), lineColor: Color.Black);
        wb.GetThemeXml().Should().BeNull();
    }

    // ---- Explicit SetThemeXml wins, before or after ------------------------

    private static readonly byte[] TinyTheme = System.Text.Encoding.UTF8.GetBytes(
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Tiny"><a:themeElements><a:clrScheme name="Tiny"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FEFEFE"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2><a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4><a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6><a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>""");

    [Fact]
    public void Explicit_Theme_Set_Before_The_Write_Is_Kept()
    {
        using var wb = Workbook.Create();
        wb.SetThemeXml(TinyTheme);
        wb.AddSheet("S")["A1"].Style(new CellStyle { FontColorTheme = new ThemeColor(4) });
        wb.ResolveThemeColor(4)!.Value.ToHex().Should().Be("#FFFF0000", "the explicit theme wins over the lazy default");
        System.Text.Encoding.UTF8.GetString(wb.GetThemeXml()!).Should().Contain("name=\"Tiny\"");
    }

    [Fact]
    public void Explicit_Theme_Set_After_The_Write_Replaces_The_Default()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].Style(new CellStyle { FontColorTheme = new ThemeColor(4) });
        wb.ResolveThemeColor(4)!.Value.ToHex().Should().Be("#FF4472C4");
        wb.SetThemeXml(TinyTheme);
        wb.ResolveThemeColor(4)!.Value.ToHex().Should().Be("#FFFF0000", "a later SetThemeXml replaces the embedded default");
    }

    // ---- Opened workbooks ---------------------------------------------------

    [Fact]
    public void Opened_File_With_A_Theme_Is_Never_Clobbered()
    {
        var path = Path.Combine(Path.GetTempPath(), $"i89-keep-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S");
                wb.SetThemeXml(TinyTheme);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["S"]["A1"].Style(new CellStyle { FontColorTheme = new ThemeColor(4) });
                wb.ResolveThemeColor(4)!.Value.ToHex().Should().Be("#FFFF0000",
                    "an existing theme part must survive new theme-indexed writes");
            }
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Opened_ThemeFree_File_Gets_The_Embed_On_A_New_Theme_Write()
    {
        var path = Path.Combine(Path.GetTempPath(), $"i89-embed-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S")["A1"].SetString("plain");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb.ResolveThemeColor(4).Should().BeNull();
                wb["S"]["A2"].Style(new CellStyle { FontColorTheme = new ThemeColor(4) });
                AssertEmbedded(wb);
            }
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void A_ZeroLength_Theme_Part_Is_Repaired_By_The_Embed()
    {
        // A part with no bytes matches GetThemeXml()'s null contract (it is
        // not a usable theme), so the first theme write repairs it in place
        // instead of resolving against nothing. Constructed via the escape
        // hatch — only a malformed producer gets here.
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Underlying.WorkbookPart!.AddNewPart<DocumentFormat.OpenXml.Packaging.ThemePart>();
        wb.GetThemeXml().Should().BeNull("a zero-length part reads as no theme");

        wb["S"]["A1"].Style(new CellStyle { FontColorTheme = new ThemeColor(4) });
        wb.ResolveThemeColor(4)!.Value.ToHex().Should().Be("#FF4472C4");
    }

    // ---- Streaming engine (assembly-time check) -----------------------------

    [Fact]
    public void Streaming_Theme_Style_Embeds_At_Assembly()
    {
        var path = Path.Combine(Path.GetTempPath(), $"i89-stream-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var swb = Workbook.CreateStreaming())
            {
                var sheet = swb.AddSheet("Big");
                var row = sheet.AppendRow();
                row.Set(1, "themed");
                row.Cell(1).Style(new CellStyle { FontColorTheme = new ThemeColor(4) });
                swb.Save(path);
            }
            using var wb = Workbook.Open(path);
            wb.GetThemeXml().Should().NotBeNull("the streaming engine embeds at Save-time assembly");
            wb.ResolveThemeColor(4)!.Value.ToHex().Should().Be("#FF4472C4");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Streaming_Plain_Workbook_Stays_Theme_Free()
    {
        var path = Path.Combine(Path.GetTempPath(), $"i89-stream-plain-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var swb = Workbook.CreateStreaming())
            {
                var sheet = swb.AddSheet("Big");
                var row = sheet.AppendRow();
                row.Set(1, "plain");
                row.Cell(1).Style(new CellStyle { Bold = true, Background = Color.FromHex("#EEEEEE") });
                swb.Save(path);
            }
            using var wb = Workbook.Open(path);
            wb.GetThemeXml().Should().BeNull();
        }
        finally { File.Delete(path); }
    }

    // ---- Schema gate ---------------------------------------------------------

    [Fact]
    public void Embedded_Theme_And_Every_New_Axis_Validate_Clean()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].Style(new CellStyle
        {
            FontColorTheme = new ThemeColor(4, 0.4),
            BackgroundTheme = new ThemeColor(5, -0.25),
            Borders = new CellBorders(Top: BorderStyle.Thin, Bottom: BorderStyle.Double)
            {
                TopColorTheme = new ThemeColor(6),
                BottomColorTheme = new ThemeColor(7, 0.2),
            },
        });
        s["A2"].SetRichText(new RichText(
            new RichTextRun("themed", new RichTextStyle { ColorTheme = new ThemeColor(8, -0.1) })));
        var pic = s.AddPicture("B2", "D6", OnePixelPng, ImageFormat.Png);
        pic.Border = new PictureBorder { ThemeColor = new ThemeColor(9) };
        s.AddConnector(ConnectorType.Bent, "E1", "G3");
        s["H1"].SetString("a"); s["H2"].SetString("b");
        s["I1"].SetNumber(1); s["I2"].SetNumber(2);
        s.AddChart(ChartType.Pie, "J1", "P12", "H1:H2", "I1:I2");

        OpenXmlValidationGate.AssertValid(wb);
    }
}
