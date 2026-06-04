// Cookbook recipe 12 — TimeAndDuration
//
// Per docs/design.md §8.1: "Demonstrate SetTime(TimeOnly) and
// SetDuration(TimeSpan) with appropriate format strings, including
// elapsed-time formatting via [h]:mm:ss."
//
// v0.3.x date-time slice: every cell type the API supports
// (DateTime, DateOnly, TimeOnly, TimeSpan) gets a default number format
// from the workbook's lazy style cache. Callers can override via
// ICell.NumberFormat(string) / ICell.Style(CellStyle).

using System;
using System.Threading.Tasks;
using NetXlsx;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Writes a single sheet showing each date/time/duration kind with the
/// default format NetXlsx applies. Useful as both a demo and a smoke
/// test of decisions I-17 (Kind preservation), I-18 / I-19 (default
/// formats), and §7.9 (elapsed-time formatting).
/// </summary>
public static class TimeAndDuration
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Times";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet(SheetName);

        // Header
        sheet.AppendRow()
            .Set(1, "Kind")
            .Set(2, "Value")
            .Set(3, "Notes");

        // DateTime — full timestamp.
        sheet.AppendRow()
            .Set(1, "DateTime")
            .Set(2, new DateTime(2026, 5, 16, 9, 30, 15, DateTimeKind.Unspecified))
            .Set(3, "default format: yyyy-mm-dd hh:mm:ss");

        // DateOnly — date without time.
        sheet.AppendRow()
            .Set(1, "DateOnly")
            .Set(2, new DateOnly(2026, 5, 16))
            .Set(3, "default format: yyyy-mm-dd (I-19)");

        // TimeOnly — time of day, fraction-of-day under the hood.
        sheet.AppendRow()
            .Set(1, "TimeOnly")
            .Set(2, new TimeOnly(9, 30, 15))
            .Set(3, "default format: h:mm:ss");

        // TimeSpan — elapsed duration.
        sheet.AppendRow()
            .Set(1, "Duration (4h 15m)")
            .Set(2, TimeSpan.FromMinutes(255))
            .Set(3, "default format: [h]:mm:ss (§7.9)");

        // TimeSpan exceeding 24h — the elapsed format renders correctly.
        sheet.AppendRow()
            .Set(1, "Duration (26h)")
            .Set(2, TimeSpan.FromHours(26))
            .Set(3, "[h]:mm:ss renders 26:00:00, not 02:00:00");

        await wb.SaveAsync(outputPath);
    }
}
