// Golden-file test for cookbook recipe 10 (NPOIEscapeHatch).

using System;
using System.IO;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using FluentAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class NPOIEscapeHatchTests
{
    [Fact]
    public async Task Recipe_Print_Area_And_Page_Setup_Survive_Save_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-escape-{Guid.NewGuid():N}.xlsx");
        try
        {
            await NPOIEscapeHatch.Run(path);

            using var wb = Workbook.Open(path);
            var sheet = wb[NPOIEscapeHatch.SheetName];

            // Data round-tripped through the regular wrapper API.
            sheet["A2"].GetString().Should().Be("North");
            sheet["E5"].GetNumber().Should().Be(1500.0);

            // Escape-hatch artifacts round-trip via .Underlying since
            // we don't model them.
            var rawWb = wb.Underlying;
            var rawSheet = sheet.Underlying;

            int sheetIndex = rawWb.GetSheetIndex(rawSheet);
            rawWb.GetPrintArea(sheetIndex).Should().Be("PrintMe!$A$1:$E$5");

            rawSheet.PrintSetup.Landscape.Should().BeTrue();
            rawSheet.PrintSetup.FitWidth.Should().Be(1);
            rawSheet.FitToPage.Should().BeTrue();

            rawSheet.Header.Center.Should().Be("Regional sales — annual");
            rawSheet.Footer.Right.Should().Be("Page &P of &N");

            rawSheet.RepeatingRows.Should().NotBeNull();
            rawSheet.RepeatingRows.FirstRow.Should().Be(0);
            rawSheet.RepeatingRows.LastRow.Should().Be(0);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
