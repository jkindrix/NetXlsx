// Cookbook recipe — SortingRanges (post-v1.1 / decision I-72).
//
// Sort a block of data rows in place with one or more sort keys.
// SortKey columns are 1-based absolute column indices; the sort is
// stable, and the range's first row is NOT treated as a header — pass
// the data rows only (exclude the header row from the range).

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a one-sheet roster and sorts the data rows by department
/// (ascending), then by salary (descending) within each department —
/// a two-key sort. The header row in A1 is left out of the sort range.
/// </summary>
public static class SortingRanges
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Roster";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet(SheetName);

        sh.AppendRow().Set(1, "Name").Set(2, "Dept").Set(3, "Salary");
        sh.AppendRow().Set(1, "Ada").Set(2, "Eng").Set(3, 140_000);
        sh.AppendRow().Set(1, "Bert").Set(2, "Sales").Set(3, 90_000);
        sh.AppendRow().Set(1, "Cleo").Set(2, "Eng").Set(3, 165_000);
        sh.AppendRow().Set(1, "Dane").Set(2, "Sales").Set(3, 120_000);

        // Sort the data rows (A2:C5 — header row A1 excluded) by Dept
        // ascending, then Salary descending. Column indices are absolute.
        sh.SortRange("A2:C5", SortKey.Asc(2), SortKey.Desc(3));

        await wb.SaveAsync(outputPath);
    }
}
