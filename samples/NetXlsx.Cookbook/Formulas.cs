// Cookbook recipe 6 — Formulas
//
// Per docs/design.md §8.1: "Write formulas referencing other cells;
// demonstrate that Excel computes on open."
//
// v0.7 sub-slice A shipped ICell.SetFormula / GetFormula. Per decision
// #46 / §7.8 we never pre-compute the cached value — Excel and other
// competent consumers recalculate on open.

using System.Threading.Tasks;
using NetXlsx;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a small "quarterly sales" sheet with per-row subtotal
/// formulas (=B*C), a column-sum total (=SUM), an average (=AVERAGE),
/// and a tax-line (=Total * 0.07). All formulas are written without a
/// pre-computed cached value — open the file in Excel and the cells
/// fill in.
/// </summary>
public static class Formulas
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Quarterly";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet(SheetName);

        sheet.AppendRow()
            .Set(1, "Product")
            .Set(2, "Unit price")
            .Set(3, "Units sold")
            .Set(4, "Revenue");

        // Three data rows.
        sheet.AppendRow().Set(1, "Widget").Set(2, 9.99m).Set(3, 120);
        sheet.AppendRow().Set(1, "Gadget").Set(2, 14.50m).Set(3, 80);
        sheet.AppendRow().Set(1, "Doohickey").Set(2, 4.25m).Set(3, 300);

        // Per-row revenue: =B2*C2, =B3*C3, =B4*C4. Both leading-=
        // and bare-body forms work — the recipe uses bare for terseness.
        sheet["D2"].SetFormula("B2*C2");
        sheet["D3"].SetFormula("B3*C3");
        sheet["D4"].SetFormula("B4*C4");

        // Summary block below.
        sheet.AppendRow();   // blank spacer row 5

        sheet["A6"].SetString("Total");
        sheet["D6"].SetFormula("=SUM(D2:D4)");

        sheet["A7"].SetString("Average");
        sheet["D7"].SetFormula("=AVERAGE(D2:D4)");

        sheet["A8"].SetString("Tax (7%)");
        sheet["D8"].SetFormula("=D6*0.07");

        await wb.SaveAsync(outputPath);
    }
}
