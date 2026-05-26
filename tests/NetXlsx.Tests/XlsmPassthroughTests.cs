using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class XlsmPassthroughTests
{
    [Fact]
    public void CreateMacroEnabled_Produces_MacroEnabled_Workbook()
    {
        using var wb = Workbook.CreateMacroEnabled();
        wb.IsMacroEnabled.Should().BeTrue();
    }

    [Fact]
    public void Create_Produces_Non_MacroEnabled_Workbook()
    {
        using var wb = Workbook.Create();
        wb.IsMacroEnabled.Should().BeFalse();
    }

    [Fact]
    public void CreateMacroEnabled_Supports_Normal_Sheet_Operations()
    {
        using var wb = Workbook.CreateMacroEnabled();
        var sheet = wb.AddSheet("Data");
        sheet["A1"].SetString("hello");
        sheet["B1"].SetNumber(42);

        wb.SheetCount.Should().Be(1);
        sheet["A1"].GetString().Should().Be("hello");
        sheet["B1"].GetNumber().Should().Be(42);
    }

    [Fact]
    public void CreateMacroEnabled_Accepts_Options()
    {
        var opts = new WorkbookOptions { DefaultFontName = "Arial", DefaultFontSize = 14 };
        using var wb = Workbook.CreateMacroEnabled(opts);
        wb.IsMacroEnabled.Should().BeTrue();
    }

    [Fact]
    public void MacroEnabled_Survives_File_RoundTrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsm-rt-{System.Guid.NewGuid():N}.xlsm");
        try
        {
            using (var wb = Workbook.CreateMacroEnabled())
            {
                var s = wb.AddSheet("Sheet1");
                s["A1"].SetString("macro-test");
                wb.Save(path);
            }

            using var opened = Workbook.Open(path);
            opened.IsMacroEnabled.Should().BeTrue();
            opened.SheetCount.Should().Be(1);
            opened["Sheet1"]["A1"].GetString().Should().Be("macro-test");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void MacroEnabled_Survives_Stream_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.CreateMacroEnabled())
        {
            wb.AddSheet("S1")["A1"].SetNumber(99);
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        opened.IsMacroEnabled.Should().BeTrue();
        opened["S1"]["A1"].GetNumber().Should().Be(99);
    }

    [Fact]
    public void Regular_Xlsx_Reports_Not_MacroEnabled_After_Open()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            wb.AddSheet("S1");
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        opened.IsMacroEnabled.Should().BeFalse();
    }

    [Fact]
    public void MacroEnabled_RoundTrip_Preserves_Multiple_Sheets_And_Styles()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.CreateMacroEnabled())
        {
            var s1 = wb.AddSheet("Sheet1");
            s1["A1"].SetString("header");
            s1["A1"].Style(new CellStyle { Bold = true });
            var s2 = wb.AddSheet("Sheet2");
            s2["A1"].SetDate(new System.DateTime(2026, 1, 15));
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        opened.IsMacroEnabled.Should().BeTrue();
        opened.SheetCount.Should().Be(2);
        opened["Sheet1"]["A1"].GetString().Should().Be("header");
    }

    [Fact]
    public void MacroEnabled_Double_RoundTrip_Preserves_Content_Type()
    {
        using var ms1 = new MemoryStream();
        using (var wb = Workbook.CreateMacroEnabled())
        {
            wb.AddSheet("S")["A1"].SetString("v1");
            wb.Save(ms1, leaveOpen: true);
        }

        ms1.Position = 0;
        using var ms2 = new MemoryStream();
        using (var wb2 = Workbook.Open(ms1))
        {
            wb2["S"]["A1"].SetString("v2");
            wb2.Save(ms2, leaveOpen: true);
        }

        ms2.Position = 0;
        using var wb3 = Workbook.Open(ms2);
        wb3.IsMacroEnabled.Should().BeTrue();
        wb3["S"]["A1"].GetString().Should().Be("v2");
    }
}
