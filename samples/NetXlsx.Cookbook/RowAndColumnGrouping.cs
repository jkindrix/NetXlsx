// Cookbook recipe — RowAndColumnGrouping (post-v1.1 / decision I-71).
//
// Excel's outline / "Group" feature. GroupRows and GroupColumns take
// 1-based inclusive ranges; nesting one group inside another increments
// the outline level (up to 7). SetRowGroupCollapsed collapses a group
// under its summary row so the workbook opens with the detail hidden.

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a one-sheet budget outline: rows 2–4 are a detail group under
/// the "Q1 total" summary row 5, with a nested sub-group on rows 2–3 to
/// show the outline level incrementing; the inner group is collapsed.
/// Columns 2–4 (the monthly columns) are grouped too.
/// </summary>
public static class RowAndColumnGrouping
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Budget";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet(SheetName);

        sh.AppendRow().Set(1, "Line").Set(2, "Jan").Set(3, "Feb").Set(4, "Mar");
        sh.AppendRow().Set(1, "Salaries").Set(2, 40).Set(3, 41).Set(4, 42);
        sh.AppendRow().Set(1, "Travel").Set(2, 5).Set(3, 6).Set(4, 4);
        sh.AppendRow().Set(1, "Supplies").Set(2, 3).Set(3, 2).Set(4, 3);
        sh.AppendRow().Set(1, "Q1 total").Set(2, 48).Set(3, 49).Set(4, 49);

        // Outer detail group: rows 2–4 roll up into the summary row 5.
        sh.GroupRows(2, 4);
        // Nested inner group: rows 2–3 sit one level deeper.
        sh.GroupRows(2, 3);
        // Open the workbook with the inner group collapsed under row 4.
        sh.SetRowGroupCollapsed(4, collapsed: true);

        // Group the three monthly columns (B–D) into one outline.
        sh.GroupColumns(2, 4);

        await wb.SaveAsync(outputPath);
    }
}
