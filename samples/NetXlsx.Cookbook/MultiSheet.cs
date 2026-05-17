// Cookbook recipe 7 — MultiSheet
//
// Per docs/design.md §8.1: "Three sheets (summary, data, lookup) with
// named ranges and cross-sheet formulas."
//
// Unblocked by v0.7 sub-slice B (IWorkbook.AddNamedRange) and v0.7
// sub-slice A (ICell.SetFormula).

using System.Threading.Tasks;
using NetXlsx;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a three-sheet workbook with named ranges and cross-sheet
/// references:
/// <list type="bullet">
///   <item><c>Data</c> — raw monthly figures.</item>
///   <item><c>Lookup</c> — a region-name lookup table.</item>
///   <item><c>Summary</c> — formulas that aggregate Data via named
///   ranges declared at workbook scope.</item>
/// </list>
/// Demonstrates the value of named ranges as documentation: the
/// summary formula reads <c>=SUM(MonthlySales)</c> rather than
/// <c>=SUM(Data!B2:B13)</c>.
/// </summary>
public static class MultiSheet
{
    /// <summary>Output sheet names.</summary>
    public const string DataSheet = "Data";
    /// <summary>Output sheet names.</summary>
    public const string LookupSheet = "Lookup";
    /// <summary>Output sheet names.</summary>
    public const string SummarySheet = "Summary";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();

        // ---- Data sheet: 12 months × (month name, sales) -------------
        var data = wb.AddSheet(DataSheet);
        data.AppendRow().Set(1, "Month").Set(2, "Sales").Set(3, "Region");

        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                             "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
        var sales  = new decimal[] { 1200, 1500, 1700, 1600, 1800, 2100,
                                     2400, 2200, 1900, 1750, 1850, 2300 };
        var region = new[] { "N", "N", "N", "S", "S", "S",
                             "E", "E", "E", "W", "W", "W" };

        for (int i = 0; i < months.Length; i++)
            data.AppendRow().Set(1, months[i]).Set(2, sales[i]).Set(3, region[i]);

        // ---- Lookup sheet: region code → display name ----------------
        var lookup = wb.AddSheet(LookupSheet);
        lookup.AppendRow().Set(1, "Code").Set(2, "Name");
        lookup.AppendRow().Set(1, "N").Set(2, "North");
        lookup.AppendRow().Set(1, "S").Set(2, "South");
        lookup.AppendRow().Set(1, "E").Set(2, "East");
        lookup.AppendRow().Set(1, "W").Set(2, "West");

        // ---- Named ranges (workbook-scoped) --------------------------
        wb.AddNamedRange("MonthlySales", $"{DataSheet}!$B$2:$B$13");
        wb.AddNamedRange("RegionLookup", $"{LookupSheet}!$A$2:$B$5");

        // ---- Summary sheet: aggregates via named ranges --------------
        var summary = wb.AddSheet(SummarySheet);
        summary.AppendRow().Set(1, "Metric").Set(2, "Value");

        summary.AppendRow().Set(1, "Total sales");
        summary["B2"].SetFormula("=SUM(MonthlySales)");

        summary.AppendRow().Set(1, "Average monthly sales");
        summary["B3"].SetFormula("=AVERAGE(MonthlySales)");

        summary.AppendRow().Set(1, "Peak month sales");
        summary["B4"].SetFormula("=MAX(MonthlySales)");

        summary.AppendRow().Set(1, "Region of first row");
        // VLOOKUP against the named lookup range.
        summary["B5"].SetFormula("=VLOOKUP(Data!C2, RegionLookup, 2, FALSE)");

        // Freeze the header row on each sheet for usability.
        data.FreezeRows(1);
        lookup.FreezeRows(1);
        summary.FreezeRows(1);

        await wb.SaveAsync(outputPath);
    }
}
