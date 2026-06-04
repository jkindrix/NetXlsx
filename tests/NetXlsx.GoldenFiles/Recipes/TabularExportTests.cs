// Golden-file test for cookbook recipe 2 (TabularExport).
// Uses a small dataset injected via the recipe's overload so the test
// runs fast and the assertions are exact (no synthetic 10k-row sweep).

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class TabularExportTests
{
    [Fact]
    public async Task Recipe_Writes_Header_Row_And_All_Data_Rows()
    {
        var rows = new[]
        {
            new TabularExport.SalesRow("North", 1000.50m, 0.12, true),
            new TabularExport.SalesRow("South", 2500.00m, 0.18, false),
            new TabularExport.SalesRow("East",  3700.75m, 0.22, true),
        };
        var path = Path.Combine(Path.GetTempPath(), $"golden-tabular-{Guid.NewGuid():N}.xlsx");
        try
        {
            await TabularExport.Run(path, rows);

            using var wb = Workbook.Open(path);
            var sheet = wb[TabularExport.SheetName];

            // Header row at A1..D1.
            sheet["A1"].GetString().Should().Be("Region");
            sheet["B1"].GetString().Should().Be("Revenue");
            sheet["C1"].GetString().Should().Be("Margin");
            sheet["D1"].GetString().Should().Be("Strategic");

            // Data rows 2..N.
            for (int i = 0; i < rows.Length; i++)
            {
                int r = i + 2;
                sheet[$"A{r}"].GetString().Should().Be(rows[i].Region);
                sheet[$"B{r}"].GetNumber().Should().Be((double)rows[i].Revenue);
                sheet[$"C{r}"].GetNumber().Should().BeApproximately(rows[i].Margin, 1e-12);
                sheet[$"D{r}"].GetBool().Should().Be(rows[i].Strategic);
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task Recipe_Handles_Empty_Dataset_With_Header_Only()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-tabular-empty-{Guid.NewGuid():N}.xlsx");
        try
        {
            await TabularExport.Run(path, new List<TabularExport.SalesRow>());

            using var wb = Workbook.Open(path);
            var sheet = wb[TabularExport.SheetName];
            sheet["A1"].GetString().Should().Be("Region");
            sheet["A2"].Kind.Should().Be(CellKind.Empty);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task Recipe_Freezes_The_Header_Row()
    {
        // v0.6 sub-slice A: TabularExport now satisfies the design's
        // originally-specced "freeze the header" cookbook requirement.
        var path = Path.Combine(Path.GetTempPath(), $"golden-tabular-freeze-{Guid.NewGuid():N}.xlsx");
        try
        {
            await TabularExport.Run(path, new[]
            {
                new TabularExport.SalesRow("North", 1m, 0.1, true),
            });

            using var wb = Workbook.Open(path);
            var pane = NetXlsx.Tests.SavedOoxml.PartFromFile(path, "xl/worksheets/sheet1.xml").Root!
                .Element(NetXlsx.Tests.SavedOoxml.Main + "sheetViews")!
                .Element(NetXlsx.Tests.SavedOoxml.Main + "sheetView")!
                .Element(NetXlsx.Tests.SavedOoxml.Main + "pane");
            pane.Should().NotBeNull("FreezeRows(1) should produce a freeze pane");
            ((double?)pane!.Attribute("ySplit") ?? 0).Should().Be(1,
                "row 1 is frozen — header stays visible while scrolling data rows");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task Recipe_Synthetic_10k_Run_Produces_Valid_Workbook()
    {
        // Smoke test the default synthetic path. We don't assert all 10k
        // values — only the bookends — because the point is "the recipe
        // completes and produces a valid round-trippable file at scale."
        var path = Path.Combine(Path.GetTempPath(), $"golden-tabular-10k-{Guid.NewGuid():N}.xlsx");
        try
        {
            await TabularExport.Run(path);

            using var wb = Workbook.Open(path);
            var sheet = wb[TabularExport.SheetName];
            sheet["A1"].GetString().Should().Be("Region");
            // First data row.
            sheet["A2"].GetString().Should().Be("North");
            // Last data row (1-indexed, 10000 records => last at row 10001).
            sheet["A10001"].GetString().Should().NotBeNullOrEmpty();
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
