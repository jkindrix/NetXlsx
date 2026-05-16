// Cookbook recipe 5 — StyledReport
//
// Per docs/design.md §8.1: "Demonstrate the fluent style API: bold
// headers, currency formatting, conditional cell coloring."
//
// v0.4 styling slice: exercises the three primary style axes (font
// weight, number format, fill color) through the CellStyle value
// record + ICell.Style API. Conditional coloring is rule-driven
// in-recipe (not a generic conditional-formatting feature — that's
// a separate v2.0 capability per the roadmap).

using System.Collections.Generic;
using System.Threading.Tasks;
using NetXlsx;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Produces a styled tabular report: bold gray-shaded header row,
/// currency-formatted revenue column, and rows highlighted yellow
/// when the margin falls below a threshold.
/// </summary>
public static class StyledReport
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Report";

    /// <summary>Margin threshold below which a row is highlighted.</summary>
    public const double LowMarginThreshold = 0.15;

    /// <summary>Reusable styles. Sharing the value records means the pool dedupes them automatically.</summary>
    private static readonly CellStyle HeaderStyle = new()
    {
        Bold = true,
        Background = Color.LightGray,
        HorizontalAlignment = HAlign.Center,
    };

    private static readonly CellStyle CurrencyStyle = new()
    {
        NumberFormat = NumberFormats.Currency,
    };

    private static readonly CellStyle LowMarginHighlight = new()
    {
        Background = Color.Yellow,
    };

    /// <summary>A sample record type.</summary>
    public sealed record ReportRow(string Region, decimal Revenue, double Margin);

    /// <summary>Runs the recipe with a small fixed dataset.</summary>
    public static Task Run(string outputPath) =>
        Run(outputPath, new[]
        {
            new ReportRow("North",  1500.00m, 0.22),
            new ReportRow("South",   850.50m, 0.08),   // below threshold -> highlighted
            new ReportRow("East",   2300.75m, 0.18),
            new ReportRow("West",    600.00m, 0.10),   // below threshold -> highlighted
            new ReportRow("Central",1100.25m, 0.15),
        });

    /// <summary>Runs the recipe with a caller-supplied dataset.</summary>
    public static async Task Run(string outputPath, IReadOnlyList<ReportRow> rows)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet(SheetName);

        // Header: bold + gray fill + centered.
        var header = sheet.AppendRow()
            .Set(1, "Region")
            .Set(2, "Revenue")
            .Set(3, "Margin");
        header.Cell(1).Style(HeaderStyle);
        header.Cell(2).Style(HeaderStyle);
        header.Cell(3).Style(HeaderStyle);

        // Data: revenue as currency; low-margin rows highlighted yellow.
        foreach (var r in rows)
        {
            var row = sheet.AppendRow()
                .Set(1, r.Region)
                .Set(2, r.Revenue)
                .Set(3, r.Margin);

            row.Cell(2).Style(CurrencyStyle);

            if (r.Margin < LowMarginThreshold)
            {
                row.Cell(1).Style(LowMarginHighlight);
                row.Cell(2).Style(LowMarginHighlight);
                row.Cell(3).Style(LowMarginHighlight);
            }
        }

        await wb.SaveAsync(outputPath);
    }
}
