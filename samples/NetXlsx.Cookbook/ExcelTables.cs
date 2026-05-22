// Cookbook recipe — ExcelTables (v1.1 / decision I-51 + I-56).
//
// A structured Excel table (ListObject) with a header row, a few data
// rows, and a built-in style. Tables auto-include AutoFilter; the
// recipe also demonstrates a standalone AutoFilter on a different
// range so both APIs are visible side-by-side.

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a workbook with one table-formatted sheet ("Sales") and one
/// range-with-AutoFilter sheet ("Inventory") so callers can compare
/// the two surfaces and pick which fits their use case.
/// </summary>
public static class ExcelTables
{
    /// <summary>Output sheet name (the table sheet).</summary>
    public const string SheetName = "Sales";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();

        // ---- Sheet 1: structured table ----
        var sales = wb.AddSheet(SheetName);
        sales.AppendRow().Set(1, "Region").Set(2, "Quarter").Set(3, "Revenue");
        sales.AppendRow().Set(1, "EU").Set(2, "Q1").Set(3, 110_000);
        sales.AppendRow().Set(1, "EU").Set(2, "Q2").Set(3, 125_000);
        sales.AppendRow().Set(1, "US").Set(2, "Q1").Set(3, 200_000);
        sales.AppendRow().Set(1, "US").Set(2, "Q2").Set(3, 215_000);

        sales.AddTable("A1:C5", "QuarterlySales", TableStyles.Medium2);

        // ---- Sheet 2: range with standalone AutoFilter (no table) ----
        var inv = wb.AddSheet("Inventory");
        inv.AppendRow().Set(1, "SKU").Set(2, "OnHand");
        inv.AppendRow().Set(1, "A-100").Set(2, 42);
        inv.AppendRow().Set(1, "B-200").Set(2, 17);
        inv.AppendRow().Set(1, "C-300").Set(2, 5);
        inv.SetAutoFilter("A1:B4");

        await wb.SaveAsync(outputPath);
    }
}
