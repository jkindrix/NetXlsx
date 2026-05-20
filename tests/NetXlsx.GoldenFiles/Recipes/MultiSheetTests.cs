// Golden-file test for cookbook recipe 7 (MultiSheet).

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class MultiSheetTests
{
    [Fact]
    public async Task Recipe_Writes_Three_Sheets_With_Named_Ranges_And_Cross_Sheet_Formulas()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-multi-{Guid.NewGuid():N}.xlsx");
        try
        {
            await MultiSheet.Run(path);

            using var wb = Workbook.Open(path);
            wb.SheetCount.Should().Be(3);

            // Data sheet — 12 months + header.
            var data = wb[MultiSheet.DataSheet];
            data["A1"].GetString().Should().Be("Month");
            data["A2"].GetString().Should().Be("Jan");
            data["B13"].GetNumber().Should().Be(2300.0);
            data["C2"].GetString().Should().Be("N");

            // Lookup sheet.
            var lookup = wb[MultiSheet.LookupSheet];
            lookup["B2"].GetString().Should().Be("North");

            // Named ranges round-tripped.
            var names = wb.NamedRanges.Select(n => n.Name).OrderBy(n => n, StringComparer.Ordinal).ToArray();
            names.Should().Equal("MonthlySales", "RegionLookup");

            var monthly = wb.NamedRanges.Single(n => n.Name == "MonthlySales");
            monthly.Formula.Should().Be($"{MultiSheet.DataSheet}!$B$2:$B$13");
            monthly.SheetScope.Should().BeNull();   // workbook-scoped

            // Summary formulas reference the named ranges.
            var summary = wb[MultiSheet.SummarySheet];
            summary["A2"].GetString().Should().Be("Total sales");
            summary["B2"].Kind.Should().Be(CellKind.Formula);
            summary["B2"].GetFormula().Should().Be("=SUM(MonthlySales)");
            summary["B3"].GetFormula().Should().Be("=AVERAGE(MonthlySales)");
            summary["B4"].GetFormula().Should().Be("=MAX(MonthlySales)");
            summary["B5"].GetFormula().Should().Be("=VLOOKUP(Data!C2, RegionLookup, 2, FALSE)");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
