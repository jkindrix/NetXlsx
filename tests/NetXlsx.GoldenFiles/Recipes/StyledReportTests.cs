// Golden-file test for cookbook recipe 5 (StyledReport).

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class StyledReportTests
{
    [Fact]
    public async Task Recipe_Produces_Styled_Report_With_Currency_Bold_Header_And_Yellow_Highlights()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-styled-{Guid.NewGuid():N}.xlsx");
        try
        {
            await StyledReport.Run(path);

            using var wb = Workbook.Open(path);
            var sheet = wb[StyledReport.SheetName];

            // Header row 1 — bold + light-gray fill + center-aligned.
            for (int col = 1; col <= 3; col++)
            {
                var s = sheet[1, col].GetStyle();
                s.Bold.Should().Be(true, $"header cell ({col}) should be bold");
                s.Background.Should().Be(Color.LightGray, $"header cell ({col}) should be gray");
                s.HorizontalAlignment.Should().Be(HAlign.Center);
            }

            // Currency formatting on the Revenue column.
            sheet[2, 2].GetStyle().NumberFormat.Should().Be(NumberFormats.Currency);

            // Row 3 (South, 0.08 margin) is below threshold — yellow on all 3 cells.
            sheet[3, 1].GetStyle().Background.Should().Be(Color.Yellow);
            sheet[3, 2].GetStyle().Background.Should().Be(Color.Yellow);
            sheet[3, 3].GetStyle().Background.Should().Be(Color.Yellow);

            // Row 4 (East, 0.18 margin) is ABOVE threshold — no highlight.
            sheet[4, 1].GetStyle().Background.Should().BeNull();
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task Recipe_Output_Style_Pool_Is_Reasonably_Sized()
    {
        // A small recipe should not produce a sprawling style table —
        // proves the dedup pool is working. We're not chasing the
        // smallest possible value; we just want to catch the failure
        // mode where every cell allocates a fresh ICellStyle.
        var path = Path.Combine(Path.GetTempPath(), $"golden-styled-pool-{Guid.NewGuid():N}.xlsx");
        try
        {
            await StyledReport.Run(path);
            using var wb = Workbook.Open(path);

            // Distinct ICellStyle indices used by the workbook's cells.
            var sheet = wb[StyledReport.SheetName];
            var usedIndices = Enumerable.Range(1, sheet.Underlying.LastRowNum + 1)
                .SelectMany(r => Enumerable.Range(1, 3).Select(c => sheet[r, c].Underlying.CellStyle.Index))
                .Distinct()
                .ToList();

            // Expected style buckets:
            // 1) HeaderStyle (bold + gray + center)            — 3 header cells
            // 2) CurrencyStyle on un-highlighted Revenue       — for above-threshold rows
            // 3) LowMarginHighlight on region/margin cells     — 2 rows × {Region, Margin} cells
            // 4) Merged: CurrencyStyle + LowMarginHighlight    — Revenue cells in highlighted rows
            // 5) "no style" (index 0) for region/margin cells not highlighted
            // So roughly 4-6 distinct indices, not 5 rows × 3 cols + header = 18.
            usedIndices.Count.Should().BeLessThan(10,
                "the style pool should dedupe — distinct ICellStyle indices should be few (saw {0})", usedIndices.Count);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
