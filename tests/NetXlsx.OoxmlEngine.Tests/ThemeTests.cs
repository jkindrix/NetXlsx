// I-82 engine swap — Open XML SDK engine, drawings slice: theme round-trip.
//
// Mirrors the NPOI-engine contract in NetXlsx.Tests/ThemeReadAndDrawingIteration
// Tests (the read-side theme half) so the cutover is de-risked: every assertion
// the legacy engine satisfies, the SDK engine must satisfy too. SetThemeXml /
// GetThemeXml exercise the ThemePart part-graph; the ResolveThemeColor /
// GetThemeLineWidthEmu cases prove the engine-agnostic ThemeInfo resolution is
// wired identically (slot mapping, tx/bg aliases, Excel tint, indexed widths).

using System;
using System.IO;
using System.Text;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class ThemeTests
{
    // A minimal but schema-realistic theme: a full 12-slot clrScheme (dk1=black
    // via sysClr, lt1=white, accent1=red … accent6=cyan), a fontScheme, and an
    // fmtScheme whose lnStyleLst carries the standard three Office line widths
    // (9525/25400/38100 EMU). Matches the NPOI-engine fixture byte-for-byte.
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

    // ---- Theme XML round-trip ---------------------------------------

    [Fact]
    public void GetThemeXml_Returns_Null_For_Workbook_Without_Theme()
    {
        using var wb = Workbook.CreateOoxml();
        wb.GetThemeXml().Should().BeNull();
    }

    [Fact]
    public void GetThemeXml_Returns_What_SetThemeXml_Wrote()
    {
        using var wb = Workbook.CreateOoxml();
        wb.SetThemeXml(TinyTheme);
        var back = wb.GetThemeXml();
        back.Should().NotBeNull();
        Encoding.UTF8.GetString(back!).Should().Contain("<a:dk1>");
    }

    [Fact]
    public void SetThemeXml_Twice_Reuses_The_Theme_Part()
    {
        using var wb = Workbook.CreateOoxml();
        wb.SetThemeXml(TinyTheme);
        // A second set must overwrite the existing part, not orphan a relationship.
        wb.SetThemeXml(TinyTheme);
        wb.Underlying.WorkbookPart!.ThemePart.Should().NotBeNull();
        Encoding.UTF8.GetString(wb.GetThemeXml()!).Should().Contain("<a:accent6>");
    }

    [Fact]
    public void SetThemeXml_Null_Throws()
    {
        using var wb = Workbook.CreateOoxml();
        Action act = () => wb.SetThemeXml(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Theme_Round_Trips_Through_File()
    {
        var path = Path.Combine(Path.GetTempPath(), $"theme-rt-ooxml-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                wb.AddSheet("S");
                wb.SetThemeXml(TinyTheme);
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                var back = wb.GetThemeXml();
                back.Should().NotBeNull();
                // Byte-equality (not a substring probe): pins the load-bearing "read the
                // part stream, never materialize ThemePart.Theme" property (lesson #2 /
                // SDK-quirk #12) so a future refactor to ThemePart.Theme.OuterXml can't
                // silently drift whitespace/ordering through the file round-trip.
                back.Should().Equal(TinyTheme);
                // Resolution survives the reopen — the part relationship is intact.
                wb.ResolveThemeColor("accent3").Should().Be(Color.FromRgb(0x00, 0x00, 0xFF));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- ResolveThemeColor by index, name, and ThemeColor ------------

    [Fact]
    public void ResolveThemeColor_By_Index_Honors_OOXML_Slot_Mapping()
    {
        using var wb = Workbook.CreateOoxml();
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
        using var wb = Workbook.CreateOoxml();
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
        using var wb = Workbook.CreateOoxml();
        wb.SetThemeXml(TinyTheme);
        wb.ResolveThemeColor(new ThemeColor(Index: 4)).Should().Be(Color.FromRgb(0xFF, 0x00, 0x00));
    }

    [Fact]
    public void ResolveThemeColor_Returns_Null_When_No_Theme()
    {
        using var wb = Workbook.CreateOoxml();
        wb.ResolveThemeColor(1).Should().BeNull();
        wb.ResolveThemeColor("dk1").Should().BeNull();
    }

    [Fact]
    public void ResolveThemeColor_Returns_Null_For_Unknown_Name()
    {
        using var wb = Workbook.CreateOoxml();
        wb.SetThemeXml(TinyTheme);
        wb.ResolveThemeColor("nonExistent").Should().BeNull();
    }

    [Fact]
    public void ResolveThemeColor_Tint_Zero_Is_The_Base_Color()
    {
        using var wb = Workbook.CreateOoxml();
        wb.SetThemeXml(TinyTheme);
        wb.ResolveThemeColor("accent1", 0).Should().Be(Color.FromRgb(0xFF, 0x00, 0x00));
    }

    [Fact]
    public void ResolveThemeColor_Negative_Tint_Darkens()
    {
        // Excel: tint < 0 darkens; pure red with tint -0.5 → approximately
        // (128, 0, 0). Allow a small rounding tolerance.
        using var wb = Workbook.CreateOoxml();
        wb.SetThemeXml(TinyTheme);
        var c = wb.ResolveThemeColor("accent1", -0.5);
        c.Should().NotBeNull();
        c!.Value.R.Should().BeInRange(120, 136);
        c.Value.G.Should().Be(0);
        c.Value.B.Should().Be(0);
    }

    [Fact]
    public void ResolveThemeColor_Re_Set_Theme_Is_Seen()
    {
        // The cached ThemeInfo must be invalidated by a second SetThemeXml so the
        // new scheme is resolved, not the stale one.
        using var wb = Workbook.CreateOoxml();
        wb.SetThemeXml(TinyTheme);
        wb.ResolveThemeColor("accent1").Should().Be(Color.FromRgb(0xFF, 0x00, 0x00));

        // A second theme where accent1 is green; resolution must reflect it.
        var greenAccent = Encoding.UTF8.GetString(TinyTheme)
            .Replace("<a:accent1><a:srgbClr val=\"FF0000\"/></a:accent1>",
                     "<a:accent1><a:srgbClr val=\"00FF00\"/></a:accent1>");
        wb.SetThemeXml(Encoding.UTF8.GetBytes(greenAccent));
        wb.ResolveThemeColor("accent1").Should().Be(Color.FromRgb(0x00, 0xFF, 0x00));
    }

    // ---- GetThemeLineWidthEmu ----------------------------------------

    [Fact]
    public void GetThemeLineWidthEmu_Reads_Indexed_Widths()
    {
        using var wb = Workbook.CreateOoxml();
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
        using var wb = Workbook.CreateOoxml();
        wb.GetThemeLineWidthEmu(1).Should().BeNull();
    }

    // ---- Disposed guard ----------------------------------------------

    [Fact]
    public void Theme_Members_Throw_After_Dispose()
    {
        var wb = Workbook.CreateOoxml();
        wb.Dispose();
        ((Action)(() => wb.GetThemeXml())).Should().Throw<ObjectDisposedException>();
        ((Action)(() => wb.SetThemeXml(TinyTheme))).Should().Throw<ObjectDisposedException>();
        ((Action)(() => wb.ResolveThemeColor(1))).Should().Throw<ObjectDisposedException>();
        ((Action)(() => wb.ResolveThemeColor("dk1"))).Should().Throw<ObjectDisposedException>();
        ((Action)(() => wb.GetThemeLineWidthEmu(1))).Should().Throw<ObjectDisposedException>();
    }
}
