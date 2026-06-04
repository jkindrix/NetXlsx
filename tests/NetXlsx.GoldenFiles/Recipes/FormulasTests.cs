// Golden-file test for cookbook recipe 6 (Formulas).

using System;
using System.IO;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class FormulasTests
{
    [Fact]
    public async Task Recipe_Writes_Formula_Cells_Without_Pre_Computed_Values()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-formulas-{Guid.NewGuid():N}.xlsx");
        try
        {
            await Formulas.Run(path);

            using var wb = Workbook.Open(path);
            var sheet = wb[Formulas.SheetName];

            // Header + 3 data rows.
            sheet["A1"].GetString().Should().Be("Product");
            sheet["A2"].GetString().Should().Be("Widget");
            sheet["B2"].GetNumber().Should().Be(9.99);
            sheet["C2"].GetNumber().Should().Be(120);

            // Per-row revenue formulas — written bare, GetFormula
            // re-adds the leading '='.
            sheet["D2"].Kind.Should().Be(CellKind.Formula);
            sheet["D2"].GetFormula().Should().Be("=B2*C2");
            sheet["D3"].GetFormula().Should().Be("=B3*C3");
            sheet["D4"].GetFormula().Should().Be("=B4*C4");

            // Summary block.
            sheet["A6"].GetString().Should().Be("Total");
            sheet["D6"].GetFormula().Should().Be("=SUM(D2:D4)");

            sheet["A7"].GetString().Should().Be("Average");
            sheet["D7"].GetFormula().Should().Be("=AVERAGE(D2:D4)");

            sheet["A8"].GetString().Should().Be("Tax (7%)");
            sheet["D8"].GetFormula().Should().Be("=D6*0.07");

            // Per decision #46 / §7.8 we never pre-compute: the persisted
            // formula cells carry no meaningful cached <v> (engines differ
            // on the artifact shape — empty <v/> vs omitted). Excel
            // recalculates on open.
            var sheetXml = NetXlsx.Tests.SavedOoxml.PartFromFile(path, "xl/worksheets/sheet1.xml");
            foreach (var address in new[] { "D2", "D6" })
            {
                var v = NetXlsx.Tests.SavedOoxml.Cell(sheetXml, address)!
                    .Element(NetXlsx.Tests.SavedOoxml.Main + "v");
                (v is null || v.Value is "" or "0").Should().BeTrue(
                    $"{address} must carry no pre-computed value (#46)");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
