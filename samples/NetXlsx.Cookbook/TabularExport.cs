// Cookbook recipe 2 — TabularExport
//
// Per docs/design.md §8.1: "Write 10k records from a list, with a header
// row, frozen header, and column widths."
//
// **v0.2.0 caveat**: the full recipe needs row iteration (IRow), freeze
// panes (ISheet.FreezePanes), and column width (IColumn.Width) — none of
// which exist in v0.2.0. This implementation uses only the v0.2.0 cell
// indexer (`sheet["A{r}"]`) and *deliberately leaves the awkwardness in
// place*. Replacing the string-interpolated address arithmetic with an
// idiomatic row API is the load-bearing motivation for the IRow slice
// (next milestone). The recipe will be rewritten there.

using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using NetXlsx;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Writes a list of records as rows of an .xlsx workbook, with a header
/// row. v0.2.0 implementation uses the cell-indexer surface; subsequent
/// milestones replace the per-cell addressing with an IRow API.
/// </summary>
public static class TabularExport
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Sales";

    /// <summary>A sample record type — public so tests can construct identical inputs.</summary>
    public sealed record SalesRow(string Region, decimal Revenue, double Margin, bool Strategic);

    /// <summary>Runs the recipe with a synthetic 10k-row dataset.</summary>
    public static Task Run(string outputPath) =>
        Run(outputPath, GenerateSyntheticRows(10_000));

    /// <summary>Runs the recipe with a caller-supplied dataset.</summary>
    public static async Task Run(string outputPath, IReadOnlyList<SalesRow> rows)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet(SheetName);

        // Header row.
        sheet["A1"].SetString("Region");
        sheet["B1"].SetString("Revenue");
        sheet["C1"].SetString("Margin");
        sheet["D1"].SetString("Strategic");

        // Data rows. With only the cell indexer available, we interpolate
        // the address per cell — clunky at scale, the point of the
        // motivating example for the IRow API slice.
        for (int i = 0; i < rows.Count; i++)
        {
            int rowNumber = i + 2;  // 1-indexed, header at row 1
            var r = rows[i];
            sheet[$"A{rowNumber}"].SetString(r.Region);
            sheet[$"B{rowNumber}"].SetNumber(r.Revenue);
            sheet[$"C{rowNumber}"].SetNumber(r.Margin);
            sheet[$"D{rowNumber}"].SetBool(r.Strategic);
        }

        await wb.SaveAsync(outputPath);
    }

    private static SalesRow[] GenerateSyntheticRows(int count)
    {
        var regions = new[] { "North", "South", "East", "West", "Central" };
        var rows = new SalesRow[count];
        for (int i = 0; i < count; i++)
        {
            var region = regions[i % regions.Length];
            decimal revenue = 1000m + (i * 7.31m);
            double margin = 0.1 + (i % 50) / 1000.0;
            bool strategic = (i % 11) == 0;
            rows[i] = new SalesRow(region, revenue, margin, strategic);
        }
        return rows;
    }

    /// <summary>Invariant string used by the test to assert deterministic generator output.</summary>
    internal static string FormatSyntheticForCheck(SalesRow r) =>
        string.Create(CultureInfo.InvariantCulture, $"{r.Region}|{r.Revenue}|{r.Margin}|{r.Strategic}");
}
