// Coverage for the I-81 read-side slice: IWorkbook.GetThemeXml /
// ResolveThemeColor / GetThemeLineWidthEmu and ISheet.Pictures /
// Connectors plus the new IPicture / IConnector property surface.

using System;
using System.IO;
using System.Linq;
using System.Text;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class ThemeReadAndDrawingIterationTests
{
    // A minimal but realistic theme: dk1=black (via sysClr), accent1=red
    // (with a 0x80 tint that we don't apply here), and three line widths
    // mirroring the standard Office theme (9525/25400/38100 EMU).
    private static readonly byte[] TinyTheme = Encoding.UTF8.GetBytes(
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Tiny">
          <a:themeElements>
            <a:clrScheme name="Tiny">
              <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
              <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
              <a:dk2><a:srgbClr val="222222"/></a:dk2>
              <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
              <a:accent1><a:srgbClr val="FF0000"/></a:accent1>
              <a:accent2><a:srgbClr val="00FF00"/></a:accent2>
              <a:accent3><a:srgbClr val="0000FF"/></a:accent3>
              <a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
              <a:accent5><a:srgbClr val="FF00FF"/></a:accent5>
              <a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
              <a:hlink><a:srgbClr val="0000EE"/></a:hlink>
              <a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
            </a:clrScheme>
            <a:fontScheme name="Tiny"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
            <a:fmtScheme name="Tiny">
              <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
              <a:lnStyleLst>
                <a:ln w="9525"/>
                <a:ln w="25400"/>
                <a:ln w="38100"/>
              </a:lnStyleLst>
              <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
              <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
            </a:fmtScheme>
          </a:themeElements>
        </a:theme>
        """);

    // 1x1 transparent PNG.
    private static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

    // ---- Theme XML round-trip ---------------------------------------

    [Fact]
    public void GetThemeXml_Returns_Null_For_Workbook_Without_Theme()
    {
        using var wb = Workbook.Create();
        wb.GetThemeXml().Should().BeNull();
    }

    [Fact]
    public void GetThemeXml_Returns_What_SetThemeXml_Wrote()
    {
        using var wb = Workbook.Create();
        wb.SetThemeXml(TinyTheme);
        var back = wb.GetThemeXml();
        back.Should().NotBeNull();
        Encoding.UTF8.GetString(back!).Should().Contain("<a:dk1>");
    }

    [Fact]
    public void Theme_Round_Trips_Through_File()
    {
        var path = Path.Combine(Path.GetTempPath(), $"theme-rt-{Guid.NewGuid():N}.xlsx");
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
                var back = wb.GetThemeXml();
                back.Should().NotBeNull();
                Encoding.UTF8.GetString(back!).Should().Contain("<a:accent3>");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- ResolveThemeColor by index, name, and ThemeColor ------------

    [Fact]
    public void ResolveThemeColor_By_Index_Honors_OOXML_Slot_Mapping()
    {
        using var wb = Workbook.Create();
        wb.SetThemeXml(TinyTheme);
        // OOXML cell-color theme index: 0=lt1, 1=dk1, 2=lt2, 3=dk2,
        // 4..9=accent1..6, 10=hlink, 11=folHlink.
        wb.ResolveThemeColor(0).Should().Be(Color.FromRgb(0xFF, 0xFF, 0xFF));
        wb.ResolveThemeColor(1).Should().Be(Color.FromRgb(0x00, 0x00, 0x00));
        wb.ResolveThemeColor(4).Should().Be(Color.FromRgb(0xFF, 0x00, 0x00));
        wb.ResolveThemeColor(9).Should().Be(Color.FromRgb(0x00, 0xFF, 0xFF));
    }

    [Fact]
    public void ResolveThemeColor_By_Name_Handles_Tx_Bg_Aliases()
    {
        using var wb = Workbook.Create();
        wb.SetThemeXml(TinyTheme);
        wb.ResolveThemeColor("dk1").Should().Be(Color.FromRgb(0, 0, 0));
        // tx1 is an alias for dk1; bg1 for lt1.
        wb.ResolveThemeColor("tx1").Should().Be(Color.FromRgb(0, 0, 0));
        wb.ResolveThemeColor("bg1").Should().Be(Color.FromRgb(0xFF, 0xFF, 0xFF));
        wb.ResolveThemeColor("accent1").Should().Be(Color.FromRgb(0xFF, 0x00, 0x00));
    }

    [Fact]
    public void ResolveThemeColor_Convenience_Overload_For_ThemeColor()
    {
        using var wb = Workbook.Create();
        wb.SetThemeXml(TinyTheme);
        wb.ResolveThemeColor(new ThemeColor(Index: 4)).Should().Be(Color.FromRgb(0xFF, 0x00, 0x00));
    }

    [Fact]
    public void ResolveThemeColor_Returns_Null_When_No_Theme()
    {
        using var wb = Workbook.Create();
        wb.ResolveThemeColor(1).Should().BeNull();
        wb.ResolveThemeColor("dk1").Should().BeNull();
    }

    [Fact]
    public void ResolveThemeColor_Returns_Null_For_Unknown_Name()
    {
        using var wb = Workbook.Create();
        wb.SetThemeXml(TinyTheme);
        wb.ResolveThemeColor("nonExistent").Should().BeNull();
    }

    [Fact]
    public void ResolveThemeColor_Tint_Zero_Is_The_Base_Color()
    {
        using var wb = Workbook.Create();
        wb.SetThemeXml(TinyTheme);
        wb.ResolveThemeColor("accent1", 0).Should().Be(Color.FromRgb(0xFF, 0x00, 0x00));
    }

    [Fact]
    public void ResolveThemeColor_Negative_Tint_Darkens()
    {
        // Excel: tint < 0 darkens; pure red with tint -0.5 → approximately
        // (128, 0, 0). Allow a small rounding tolerance.
        using var wb = Workbook.Create();
        wb.SetThemeXml(TinyTheme);
        var c = wb.ResolveThemeColor("accent1", -0.5);
        c.Should().NotBeNull();
        c!.Value.R.Should().BeInRange(120, 136);
        c.Value.G.Should().Be(0);
        c.Value.B.Should().Be(0);
    }

    // ---- GetThemeLineWidthEmu ----------------------------------------

    [Fact]
    public void GetThemeLineWidthEmu_Reads_Indexed_Widths()
    {
        using var wb = Workbook.Create();
        wb.SetThemeXml(TinyTheme);
        wb.GetThemeLineWidthEmu(1).Should().Be(9525);
        wb.GetThemeLineWidthEmu(2).Should().Be(25400);
        wb.GetThemeLineWidthEmu(3).Should().Be(38100);
        wb.GetThemeLineWidthEmu(4).Should().BeNull();
        wb.GetThemeLineWidthEmu(0).Should().BeNull();
    }

    [Fact]
    public void GetThemeLineWidthEmu_Null_When_Theme_Absent()
    {
        using var wb = Workbook.Create();
        wb.GetThemeLineWidthEmu(1).Should().BeNull();
    }

    // ---- ISheet.Pictures + IPicture properties ----------------------

    [Fact]
    public void Pictures_Enumerates_Added_Pictures_With_Anchor_And_Bytes()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddPicture("B3", "D6", OnePixelPng, ImageFormat.Png, dx1: 100, dy1: 200, dx2: 300, dy2: 400);

        var pics = s.Pictures;
        pics.Should().HaveCount(1);
        var pic = pics[0];
        pic.Sheet.Should().BeSameAs(s);
        pic.Format.Should().Be(ImageFormat.Png);
        pic.FromCell.Should().Be("B3");
        pic.ToCell.Should().Be("D6");
        pic.Dx1.Should().Be(100); pic.Dy1.Should().Be(200);
        pic.Dx2.Should().Be(300); pic.Dy2.Should().Be(400);
        pic.Data.Should().BeEquivalentTo(OnePixelPng);
    }

    [Fact]
    public void Pictures_Empty_When_No_Drawing()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S").Pictures.Should().BeEmpty();
    }

    // ---- ISheet.Connectors + IConnector properties -------------------

    [Fact]
    public void Connectors_Enumerates_Added_Connectors_With_All_Properties()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConnector(ConnectorType.Straight, "I41", "I41",
            lineColor: Color.FromRgb(0, 0, 0),
            dx1: 115661, dy1: 272142, dx2: 891268, dy2: 278946,
            flipH: true, tailEnd: ConnectorEnd.Arrow, lineWidthPoints: 2.0);

        var c = s.Connectors.Should().ContainSingle().Subject;
        c.Type.Should().Be(ConnectorType.Straight);
        c.FromCell.Should().Be("I41");
        c.ToCell.Should().Be("I41");
        c.Dx1.Should().Be(115661); c.Dy2.Should().Be(278946);
        c.FlipH.Should().BeTrue();
        c.FlipV.Should().BeFalse();
        c.HeadEnd.Should().Be(ConnectorEnd.None);
        c.TailEnd.Should().Be(ConnectorEnd.Arrow);
        c.LineColor.Should().Be(Color.FromRgb(0, 0, 0));
        c.LineWidthPoints.Should().BeApproximately(2.0, 0.001);
    }

    [Fact]
    public void Connectors_Enumerates_Multiple()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConnector(ConnectorType.Straight, "A1", "C3");
        s.AddConnector(ConnectorType.Bent, "D1", "F3");
        s.AddConnector(ConnectorType.Curved, "G1", "I3");

        s.Connectors.Select(c => c.Type).Should().Equal(
            ConnectorType.Straight, ConnectorType.Bent, ConnectorType.Curved);
    }
}
