// Cookbook recipe — AutoFilterCriteria (post-v1.1 / decisions I-66, I-68).
//
// Beyond the bare dropdown arrows (SetAutoFilter), NetXlsx writes
// per-column criteria. SetAutoFilterColumn takes a 0-based column offset
// (relative to the filter range) and a FilterCriteria built from the
// static factories — comparisons, Between, Contains, In, etc. — which
// can be paired with And/Or (at most two conditions per column).
// ExcelTables.cs covers the bare AutoFilter; this recipe adds criteria.

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a one-sheet orders list with an AutoFilter over A1:C7 and two
/// per-column criteria: keep only EU/US regions (column 0, an Or pair)
/// and only amounts of at least 100 (column 2, a numeric comparison).
/// </summary>
public static class AutoFilterCriteria
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Orders";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet(SheetName);

        sh.AppendRow().Set(1, "Region").Set(2, "SKU").Set(3, "Amount");
        sh.AppendRow().Set(1, "EU").Set(2, "A-1").Set(3, 140);
        sh.AppendRow().Set(1, "US").Set(2, "B-2").Set(3, 80);
        sh.AppendRow().Set(1, "APAC").Set(2, "C-3").Set(3, 220);
        sh.AppendRow().Set(1, "EU").Set(2, "D-4").Set(3, 95);
        sh.AppendRow().Set(1, "US").Set(2, "E-5").Set(3, 310);
        sh.AppendRow().Set(1, "EU").Set(2, "F-6").Set(3, 60);

        sh.SetAutoFilter("A1:C7");

        // Column 0 (Region): keep EU OR US.
        sh.SetAutoFilterColumn(
            0, FilterCriteria.EqualTo("EU").Or(FilterCriteria.EqualTo("US")));

        // Column 2 (Amount): keep values >= 100.
        sh.SetAutoFilterColumn(2, FilterCriteria.GreaterThanOrEqual(100));

        await wb.SaveAsync(outputPath);
    }
}
