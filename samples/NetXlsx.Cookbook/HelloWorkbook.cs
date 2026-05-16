// Cookbook recipe 1 — HelloWorkbook
//
// Per docs/design.md §8.1: "Create, add a sheet, write a few cells, save."
//
// Demonstrates the v0.2.0 vertical-slice API in its simplest end-to-end
// form. Each recipe class exposes a static `Run(string outputPath)` so
// both the cookbook executable and the golden-file tests invoke it
// identically.

using System.Threading.Tasks;
using NetXlsx;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Creates a workbook with a single sheet and a handful of cells of each
/// supported scalar kind, then saves to the given path.
/// </summary>
public static class HelloWorkbook
{
    /// <summary>Output sheet name expected by both the recipe runner and its test.</summary>
    public const string SheetName = "Hello";

    /// <summary>Runs the recipe, producing an .xlsx at <paramref name="outputPath"/>.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet(SheetName);

        sheet["A1"].SetString("Greeting");
        sheet["B1"].SetString("Hello, world");

        sheet["A2"].SetString("Answer");
        sheet["B2"].SetNumber(42.0);   // 42 alone is ambiguous between SetNumber(double) and SetNumber(decimal) —
                                       // tracked as a v0.2.0 ergonomics gap to resolve in the IRow slice
                                       // (likely via int/long overloads on ICell).

        sheet["A3"].SetString("Pi");
        sheet["B3"].SetNumber(3.14159265358979);

        sheet["A4"].SetString("Sale price");
        sheet["B4"].SetNumber(19.99m);

        sheet["A5"].SetString("Is shipped?");
        sheet["B5"].SetBool(true);

        await wb.SaveAsync(outputPath);
    }
}
