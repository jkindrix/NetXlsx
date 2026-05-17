// Cookbook recipe 9 — StreamingMillionRows
//
// Per docs/design.md §8.1: "Workbook.CreateStreaming() to write a
// million rows under the perf budget."
//
// Budget from design §5: < 30s wall time, < 200 MB ΔWS for
// 1M rows × 20 cols streaming. This recipe ships at 250k rows × 20
// cols by default (≈ 0.25× the budgeted size) so it can run inside
// CI in seconds; a Run(path, rowCount) overload lets ops bump it up
// for a real perf check.

using System;
using System.Threading.Tasks;
using NetXlsx;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Writes a tabular workbook of <c>n</c> rows × 20 columns through
/// the streaming write path. Demonstrates that
/// <see cref="Workbook.CreateStreaming"/> is the right entry point
/// for any write over ~30k rows (per spike 2 / design §5) — and that
/// it composes the same fluent <c>Set</c> overloads as the
/// random-access API.
/// </summary>
public static class StreamingMillionRows
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "BigData";

    /// <summary>Default row count: 250,000 (CI-friendly).</summary>
    public const int DefaultRowCount = 250_000;

    /// <summary>Number of columns per row.</summary>
    public const int Columns = 20;

    /// <summary>Runs the recipe at the default row count.</summary>
    public static Task Run(string outputPath) => Run(outputPath, DefaultRowCount);

    /// <summary>Runs the recipe at an explicit row count.</summary>
    public static async Task Run(string outputPath, int rowCount)
    {
        // Larger window than NPOI's default (100) — at 20 cols per row,
        // 1,000 rows in-window is still ~20k cells, well below the
        // memory budget, and reduces flush churn on big writes.
        var options = new StreamingOptions { RowAccessWindowSize = 1_000 };

        await using var wb = Workbook.CreateStreaming(options);
        var sheet = wb.AddSheet(SheetName);

        // Header row.
        var header = sheet.AppendRow();
        header.Set(1, "Id");
        for (int c = 2; c <= Columns; c++)
            header.Set(c, $"Col{c}");

        // Data rows. Mix int + double + string so this isn't just a
        // numeric-fast-path demo. Deterministic values so the golden
        // test can spot-check them.
        for (int r = 1; r <= rowCount; r++)
        {
            var row = sheet.AppendRow();
            row.Set(1, r);
            for (int c = 2; c <= Columns; c++)
            {
                if (c % 3 == 0)
                    row.Set(c, (double)(r * c) / 7.0);
                else if (c % 3 == 1)
                    row.Set(c, $"r{r}-c{c}");
                else
                    row.Set(c, r * c);
            }
        }

        await wb.SaveAsync(outputPath);
    }
}
