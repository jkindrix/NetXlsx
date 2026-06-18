// Cookbook recipe — Charts (post-v1.1 / decision I-75).
//
// A single-series chart anchored between two cells, sourced from a
// category range (labels / X-axis) and a value range (Y-axis). NetXlsx
// covers the six chart types NPOI exposes: Line, Bar, Column, Pie,
// Scatter, Area. This recipe builds a column chart and a pie chart over
// the same data so two of the six are visible.

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a one-sheet workbook with a small quarterly-revenue table and
/// two charts over it: a column chart titled "Revenue by Quarter" and a
/// pie chart titled "Revenue Share".
/// </summary>
public static class Charts
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Dashboard";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet(SheetName);

        sh.AppendRow().Set(1, "Q1").Set(2, 110_000);
        sh.AppendRow().Set(1, "Q2").Set(2, 125_000);
        sh.AppendRow().Set(1, "Q3").Set(2, 138_000);
        sh.AppendRow().Set(1, "Q4").Set(2, 152_000);

        // Column chart, anchored D1:K15, categories A1:A4, values B1:B4.
        sh.AddChart(ChartType.Column, "D1", "K15", "A1:A4", "B1:B4", "Revenue by Quarter");

        // Pie chart over the same data, anchored further down the sheet.
        sh.AddChart(ChartType.Pie, "D17", "K31", "A1:A4", "B1:B4", "Revenue Share");

        await wb.SaveAsync(outputPath);
    }
}
