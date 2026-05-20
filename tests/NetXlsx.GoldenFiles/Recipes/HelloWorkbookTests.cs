// Golden-file test for cookbook recipe 1 (HelloWorkbook).
//
// Strategy: invoke the recipe's Run method, then open the produced .xlsx
// via NetXlsx again and assert the expected content. Binary fixture
// committed-in-repo is deferred until styling/formatting features land
// (when byte-level comparison becomes meaningful) — see roadmap process
// rules on fixture provenance + spike-1 / spike-2 boundary discussion.

using System;
using System.IO;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class HelloWorkbookTests
{
    [Fact]
    public async Task Recipe_Produces_Workbook_With_Expected_Content()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-hello-{Guid.NewGuid():N}.xlsx");
        try
        {
            await HelloWorkbook.Run(path);

            // File exists and is non-trivial in size (sanity check that
            // SaveAsync actually wrote something).
            File.Exists(path).Should().BeTrue();
            new FileInfo(path).Length.Should().BeGreaterThan(1000,
                "an .xlsx with content should comfortably exceed 1 KB after OPC packaging");

            using var wb = Workbook.Open(path);
            wb.SheetCount.Should().Be(1);
            var sheet = wb[HelloWorkbook.SheetName];

            // Column A — labels (strings).
            sheet["A1"].Kind.Should().Be(CellKind.String);
            sheet["A1"].GetString().Should().Be("Greeting");
            sheet["A2"].GetString().Should().Be("Answer");
            sheet["A3"].GetString().Should().Be("Pi");
            sheet["A4"].GetString().Should().Be("Sale price");
            sheet["A5"].GetString().Should().Be("Is shipped?");

            // Column B — values of each kind.
            sheet["B1"].Kind.Should().Be(CellKind.String);
            sheet["B1"].GetString().Should().Be("Hello, world");

            sheet["B2"].Kind.Should().Be(CellKind.Number);
            sheet["B2"].GetNumber().Should().Be(42.0);

            sheet["B3"].Kind.Should().Be(CellKind.Number);
            sheet["B3"].GetNumber().Should().BeApproximately(3.14159265358979, 1e-13);

            sheet["B4"].Kind.Should().Be(CellKind.Number);
            sheet["B4"].GetNumber().Should().Be(19.99);

            sheet["B5"].Kind.Should().Be(CellKind.Bool);
            sheet["B5"].GetBool().Should().Be(true);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task Recipe_Output_Is_Reopenable_As_Plain_Xlsx_Via_NPOI_Directly()
    {
        // Doubles as a guard: if NetXlsx has a write-side bug that
        // produces files only NetXlsx can open, this test catches it
        // by routing through raw NPOI for the read.
        var path = Path.Combine(Path.GetTempPath(), $"golden-hello-npoi-{Guid.NewGuid():N}.xlsx");
        try
        {
            await HelloWorkbook.Run(path);

            using var fs = File.OpenRead(path);
            using var wb = new NPOI.XSSF.UserModel.XSSFWorkbook(fs);
            wb.NumberOfSheets.Should().Be(1);
            var sheet = wb.GetSheet(HelloWorkbook.SheetName);
            sheet.Should().NotBeNull();
            sheet.GetRow(0).GetCell(0).StringCellValue.Should().Be("Greeting");
            sheet.GetRow(0).GetCell(1).StringCellValue.Should().Be("Hello, world");
            sheet.GetRow(1).GetCell(1).NumericCellValue.Should().Be(42.0);
            sheet.GetRow(4).GetCell(1).BooleanCellValue.Should().Be(true);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
