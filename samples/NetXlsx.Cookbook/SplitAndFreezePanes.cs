// Cookbook recipe — SplitAndFreezePanes (post-v1.1 / decision I-70).
//
// Two ways to keep headers in view while scrolling. A *freeze* pane
// locks rows/columns at a fixed boundary (FreezeRows/Columns/Pane). A
// *split* pane (CreateSplitPane) gives the user a draggable divider —
// distinct from freeze, and measured in twips (1/20th of a point).

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a two-sheet workbook: "Frozen" freezes the header row and the
/// first column (FreezePane); "Split" places a draggable both-direction
/// split via CreateSplitPane so the two pane styles can be compared.
/// </summary>
public static class SplitAndFreezePanes
{
    /// <summary>The freeze-pane sheet name.</summary>
    public const string FrozenSheet = "Frozen";

    /// <summary>The split-pane sheet name.</summary>
    public const string SplitSheet = "Split";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();

        // ---- Sheet 1: freeze the top row and the leftmost column. ----
        var frozen = wb.AddSheet(FrozenSheet);
        frozen.AppendRow().Set(1, "Item").Set(2, "Jan").Set(3, "Feb").Set(4, "Mar");
        for (int r = 1; r <= 20; r++)
        {
            frozen.AppendRow().Set(1, $"SKU-{r:000}").Set(2, r * 3).Set(3, r * 5).Set(4, r * 7);
        }
        frozen.FreezePane(rows: 1, cols: 1);

        // ---- Sheet 2: a draggable split, both directions (twips). ----
        var split = wb.AddSheet(SplitSheet);
        split.AppendRow().Set(1, "Item").Set(2, "Value");
        for (int r = 1; r <= 20; r++)
        {
            split.AppendRow().Set(1, $"Row {r}").Set(2, r * 11);
        }
        split.CreateSplitPane(xSplitTwips: 2000, ySplitTwips: 3000);

        await wb.SaveAsync(outputPath);
    }
}
