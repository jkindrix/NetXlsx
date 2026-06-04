// Golden-file test for cookbook recipe 10 (OpenXmlEscapeHatch).

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using AwesomeAssertions;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.GoldenFiles.Recipes;

public class OpenXmlEscapeHatchTests
{
    [Fact]
    public async Task Recipe_Print_Area_And_Page_Setup_Survive_Save_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-escape-{Guid.NewGuid():N}.xlsx");
        try
        {
            await OpenXmlEscapeHatch.Run(path);

            using var wb = Workbook.Open(path);
            var sheet = wb[OpenXmlEscapeHatch.SheetName];

            // Data round-tripped through the regular wrapper API.
            sheet["A2"].GetString().Should().Be("North");
            sheet["E5"].GetNumber().Should().Be(1500.0);

            // Escape-hatch artifacts round-trip via .Underlying since
            // we don't model them.
            var workbook = wb.Underlying.WorkbookPart!.Workbook!;
            var worksheet = sheet.Underlying;

            var names = workbook.GetFirstChild<S.DefinedNames>()!
                .Elements<S.DefinedName>().ToDictionary(n => n.Name!.Value!, n => n.Text);
            names["_xlnm.Print_Area"].Should().Be("PrintMe!$A$1:$E$5");
            names["_xlnm.Print_Titles"].Should().Be("PrintMe!$1:$1");

            var pageSetup = worksheet.GetFirstChild<S.PageSetup>()!;
            pageSetup.Orientation!.Value.Should().Be(S.OrientationValues.Landscape);
            pageSetup.FitToWidth!.Value.Should().Be(1u);
            pageSetup.FitToHeight!.Value.Should().Be(0u);
            worksheet.GetFirstChild<S.SheetProperties>()!
                .GetFirstChild<S.PageSetupProperties>()!.FitToPage!.Value.Should().BeTrue();

            var headerFooter = worksheet.GetFirstChild<S.HeaderFooter>()!;
            headerFooter.OddHeader!.Text.Should().Be("&CRegional sales — annual");
            headerFooter.OddFooter!.Text.Should().Be("&RPage &P of &N");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
