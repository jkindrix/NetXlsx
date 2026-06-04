using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class DefaultColumnWidthTests
{
    [Fact]
    public void New_Sheet_Has_Null_DefaultColumnWidth()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.DefaultColumnWidth.Should().BeNull();
    }

    [Fact]
    public void DefaultColumnWidth_Can_Be_Set_And_Read()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.DefaultColumnWidth = 12.5;
        s.DefaultColumnWidth.Should().Be(12.5);
    }

    [Fact]
    public void DefaultColumnWidth_Null_Suppresses_XML_Attribute()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");

        // The persisted sheetFormatPr must omit @defaultColWidth so Excel
        // derives the width from the Normal font metrics (I-78).
        var sfp = SavedOoxml.SheetXml(wb).Root!
            .Element(SavedOoxml.Main + "sheetFormatPr");
        ((double?)sfp?.Attribute("defaultColWidth"))
            .Should().BeNull("an unset default width must be omitted from the XML");
    }

    [Fact]
    public void Setting_DefaultColumnWidth_To_Null_Clears_It()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.DefaultColumnWidth = 10.0;
        s.DefaultColumnWidth.Should().Be(10.0);
        s.DefaultColumnWidth = null;
        s.DefaultColumnWidth.Should().BeNull();
    }

    [Fact]
    public void Normal_CellStyle_Exists_In_New_Workbook()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            wb.AddSheet("S");
            wb.Save(ms, leaveOpen: true);
        }

        // Verify by re-opening with NPOI directly
        ms.Position = 0;
        var npoi = new NPOI.XSSF.UserModel.XSSFWorkbook(ms);
        var styles = npoi.GetStylesSource();
        var doc = styles.GetType()
            .GetField("doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(styles)!;
        var ct = doc.GetType().GetMethod("GetStyleSheet")!
            .Invoke(doc, null) as NPOI.OpenXmlFormats.Spreadsheet.CT_Stylesheet;

        ct.Should().NotBeNull();
        ct!.cellStyles.Should().NotBeNull();
        ct.cellStyles.cellStyle.Should().Contain(
            cs => cs.name == "Normal" && cs.builtinId == 0,
            "the Normal built-in cellStyle must exist for Excel to resolve font metrics correctly");
        npoi.Close();
    }
}

public class ThemeColorTests
{
    [Fact]
    public void BackgroundTheme_Applies_Through_Style_Merge()
    {
        using var ms = new System.IO.MemoryStream();
        using (var wb = Workbook.Create(new WorkbookOptions { DefaultFontName = "Calibri" }))
        {
            var s = wb.AddSheet("S");
            s[1, 1].SetString("x");
            s[1, 1].Style(new CellStyle { BackgroundTheme = new ThemeColor(3, 0.4) });
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        var npoi = new NPOI.XSSF.UserModel.XSSFWorkbook(ms);
        var cs = (NPOI.XSSF.UserModel.XSSFCellStyle)
            ((NPOI.XSSF.UserModel.XSSFSheet)npoi.GetSheetAt(0)).GetRow(0).GetCell(0).CellStyle;
        cs.FillPattern.Should().Be(NPOI.SS.UserModel.FillPattern.SolidForeground);
        var fg = cs.FillForegroundXSSFColor;
        fg.Should().NotBeNull();
        fg!.Theme.Should().Be(3);
        fg.Tint.Should().BeApproximately(0.4, 0.001);
        npoi.Close();
    }
}
