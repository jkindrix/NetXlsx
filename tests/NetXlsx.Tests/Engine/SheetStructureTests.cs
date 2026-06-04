// I-82 engine swap — structure slice (5b): sheet visibility, gridlines, and
// default column width conformance. Mirrors the NPOI engine's ISheet.Hidden /
// ShowGridlines / DefaultColumnWidth contract on the Open XML SDK engine —
// visibility on workbook.xml <sheet @state>, gridlines on <sheetView
// @showGridLines> (default true), default width on <sheetFormatPr @defaultColWidth>
// (lesson #3 / I-78). Cross-checked against NetXlsx.Tests.FreezeMergeHiddenTests.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.Tests.Engine;

public class SheetStructureTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-struct-{Guid.NewGuid():N}.xlsx");

    // ---- Sheet visibility ---------------------------------------------------

    [Fact]
    public void Hidden_Defaults_False()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S").Hidden.Should().BeFalse();
    }

    [Fact]
    public void Hidden_Set_True_Writes_State_On_The_Workbook_Sheet_Element()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Visible");
        var hide = wb.AddSheet("HideMe");
        hide.Hidden = true;

        hide.Hidden.Should().BeTrue();
        // The state lives on workbook.xml <sheet>, not on the worksheet.
        var sheetEl = wb.Underlying.WorkbookPart!.Workbook!.GetFirstChild<S.Sheets>()!
            .Elements<S.Sheet>().Single(x => x.Name == "HideMe");
        sheetEl.State!.Value.Should().Be(S.SheetStateValues.Hidden);
    }

    [Fact]
    public void Hidden_Set_False_Clears_The_State_Attribute()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Visible");
        var s = wb.AddSheet("S");
        s.Hidden = true;
        s.Hidden = false;

        s.Hidden.Should().BeFalse();
        var sheetEl = wb.Underlying.WorkbookPart!.Workbook!.GetFirstChild<S.Sheets>()!
            .Elements<S.Sheet>().Single(x => x.Name == "S");
        sheetEl.State.Should().BeNull("visible is the default — no @state attribute");
    }

    [Fact]
    public void Hidden_RoundTrips_Through_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("Visible");
                wb.AddSheet("HideMe").Hidden = true;
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["Visible"].Hidden.Should().BeFalse();
                wb["HideMe"].Hidden.Should().BeTrue();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Gridlines ----------------------------------------------------------

    [Fact]
    public void ShowGridlines_Defaults_True()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S").ShowGridlines.Should().BeTrue();
    }

    [Fact]
    public void ShowGridlines_Toggles_And_RoundTrips()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddSheet("S");
                s.ShowGridlines = false;
                s.ShowGridlines.Should().BeFalse();
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["S"].ShowGridlines.Should().BeFalse();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void ShowGridlines_Set_Back_To_True_Clears_The_Attribute()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.ShowGridlines = false;
        s.ShowGridlines = true;

        s.ShowGridlines.Should().BeTrue();
        var view = wb.Underlying.WorkbookPart!.WorksheetParts.Single().Worksheet!
            .GetFirstChild<S.SheetViews>()!.GetFirstChild<S.SheetView>()!;
        view.ShowGridLines.Should().BeNull("true is the default — no @showGridLines attribute");
    }

    // ---- Default column width -----------------------------------------------

    [Fact]
    public void DefaultColumnWidth_Defaults_Null()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S").DefaultColumnWidth.Should().BeNull();
    }

    [Fact]
    public void DefaultColumnWidth_Set_And_RoundTrips()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddSheet("S");
                s.DefaultColumnWidth = 18.5;
                s.DefaultColumnWidth.Should().Be(18.5);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["S"].DefaultColumnWidth.Should().Be(18.5);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void DefaultColumnWidth_Set_Null_Clears_It()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.DefaultColumnWidth = 20;
        s.DefaultColumnWidth = null;
        s.DefaultColumnWidth.Should().BeNull();
    }

    [Fact]
    public void DefaultColumnWidth_Creates_A_Schema_Valid_SheetFormatPr()
    {
        // <sheetFormatPr> requires @defaultRowHeight — setting only the column
        // width must still produce a schema-valid element.
        using var wb = Workbook.Create();
        wb.AddSheet("S").DefaultColumnWidth = 22;
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Visibility_Gridlines_And_Width_Are_Schema_Valid()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Visible");
        var s = wb.AddSheet("S");
        s.Hidden = true;
        s.ShowGridlines = false;
        s.DefaultColumnWidth = 16;
        OpenXmlValidationGate.AssertValid(wb);
    }
}
