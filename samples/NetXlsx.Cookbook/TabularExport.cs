// Cookbook recipe 2 — TabularExport
//
// Per docs/design.md §8.1: "Write 10k records from a list, with a header
// row, frozen header, and column widths."
//
// v0.3.x: rewritten to use the IRow API (ISheet.AppendRow, IRow.Set
// fluent setters). Compare with the v0.2.0 version (git log -p) to see
// the ergonomic delta — per-cell string-interpolated addressing is
// replaced by one AppendRow per record and chained Set calls keyed by
// column index. Freeze panes and column widths still wait for a later
// slice (they require ISheet.FreezePanes and IColumn.Width).

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

        // Header row — fluent across one row.
        sheet.AppendRow()
            .Set(1, "Region")
            .Set(2, "Revenue")
            .Set(3, "Margin")
            .Set(4, "Strategic");

        // Data rows — one AppendRow per record, chained Set calls keyed
        // by column index. No string interpolation, no row arithmetic.
        foreach (var r in rows)
        {
            sheet.AppendRow()
                .Set(1, r.Region)
                .Set(2, r.Revenue)
                .Set(3, r.Margin)
                .Set(4, r.Strategic);
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
