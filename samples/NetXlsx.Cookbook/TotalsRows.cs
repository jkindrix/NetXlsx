// Cookbook recipe — TotalsRows (post-v1.1 / decision I-64).
//
// An Excel table can carry a totals row. AddTotalsRow extends the
// table's range by one row; SetColumnTotal then assigns a built-in
// aggregate (TotalsRowFunction.Sum / Average / Count / …) per column,
// which serializes as a 100-series SUBTOTAL formula. A label can be set
// on a non-aggregated column. SetColumnTotal requires AddTotalsRow first.

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a one-sheet sales table with a totals row: the "Region" column
/// gets a "Total" label, "Units" is summed, and "Revenue" is averaged.
/// </summary>
public static class TotalsRows
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Sales";

    /// <summary>The table name.</summary>
    public const string TableName = "RegionSales";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet(SheetName);

        sh.AppendRow().Set(1, "Region").Set(2, "Units").Set(3, "Revenue");
        sh.AppendRow().Set(1, "EU").Set(2, 120).Set(3, 110_000);
        sh.AppendRow().Set(1, "US").Set(2, 200).Set(3, 215_000);
        sh.AppendRow().Set(1, "APAC").Set(2, 90).Set(3, 88_000);

        var table = sh.AddTable("A1:C4", TableName, TableStyles.Medium2);

        // Add the totals row, then assign per-column aggregates by column
        // name (the header text). Region gets a plain label.
        table.AddTotalsRow();
        table.SetColumnTotalLabel("Region", "Total");
        table.SetColumnTotal("Units", TotalsRowFunction.Sum);
        table.SetColumnTotal("Revenue", TotalsRowFunction.Average);

        await wb.SaveAsync(outputPath);
    }
}
