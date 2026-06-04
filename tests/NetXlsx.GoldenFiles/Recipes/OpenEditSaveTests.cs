// Golden-file test for cookbook recipe 11 (OpenEditSave).
//
// Asserts both preservation promises from design §7.5 and §7.7:
//   §7.5 — re-applying an identical style doesn't allocate a new
//          ICellStyle (style pool dedup).
//   §7.7 — OPC parts NetXlsx doesn't model round-trip verbatim.

using System;
using System.IO;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class OpenEditSaveTests
{
    [Fact]
    public async Task Recipe_Mutations_Land_And_Unmodeled_Parts_Survive()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-oes-{Guid.NewGuid():N}.xlsx");
        try
        {
            await OpenEditSave.Run(path);

            using (var wb = Workbook.Open(path))
            {
                var sheet = wb[OpenEditSave.SheetName];

                // Mutations from the recipe.
                sheet["A4"].GetString().Should().Be("East");
                sheet["B4"].GetNumber().Should().Be(1450.0);
                sheet["B3"].GetNumber().Should().Be(1175.0);

                // §7.5 — identical CellStyle applied to A1 and B1 yields
                // the same persisted style index (pool dedup).
                var sheetXml = NetXlsx.Tests.SavedOoxml.PartFromFile(path, "xl/worksheets/sheet1.xml");
                var idxA1 = NetXlsx.Tests.SavedOoxml.CellStyleIndex(sheetXml, "A1");
                var idxB1 = NetXlsx.Tests.SavedOoxml.CellStyleIndex(sheetXml, "B1");
                idxA1.Should().NotBeNull("styled cells must carry an explicit style index");
                idxA1.Should().Be(idxB1,
                    "§7.5 — identical CellStyle records should share one style via the pool");

                sheet["A1"].GetStyle().Bold.Should().Be(true);
            }

            // §7.7 — the custom OPC part the input file carried should
            // be present in the output file with identical bytes.
            var preserved = OpenEditSave.ReadCustomPartBytes(path);
            preserved.Should().NotBeNull("§7.7 — unmodeled OPC parts must round-trip verbatim");
            preserved.Should().Equal(OpenEditSave.ExpectedCustomPartBytes);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
