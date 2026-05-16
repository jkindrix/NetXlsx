// Cookbook recipe 4 — TypedImport
//
// Per docs/design.md §8.1: "Read a workbook into a typed record sequence."
//
// v0.5.x source-gen read-side slice: demonstrates the inverse of
// TypedExport. The same [Worksheet]-annotated SalesRecord type (defined
// in TypedExport.cs) gets a sibling ReadRows extension method that
// resolves [Column(Name)] -> header column index and yields records.

using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using NetXlsx;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Round-trip recipe: write a typed dataset (via <see cref="TypedExport"/>'s
/// path), reopen the file, and read it back as a typed sequence.
/// </summary>
public static class TypedImport
{
    /// <summary>Output sheet name (matches TypedExport's).</summary>
    public const string SheetName = TypedExport.SheetName;

    /// <summary>Runs the recipe with a small fixed dataset; discards the result.</summary>
    public static async Task Run(string outputPath)
    {
        var records = new[]
        {
            new SalesRecord { Region = "North", Revenue = 1000.50m, Margin = 0.12, Strategic = true  },
            new SalesRecord { Region = "South", Revenue = 2500.00m, Margin = 0.18, Strategic = false },
            new SalesRecord { Region = "East",  Revenue = 3700.75m, Margin = 0.22, Strategic = true  },
        };
        _ = await RunAndReturn(outputPath, records);
    }

    /// <summary>Writes the dataset, then reads it back; returns the parsed records.</summary>
    public static async Task<IReadOnlyList<SalesRecord>> RunAndReturn(string outputPath, IReadOnlyList<SalesRecord> records)
    {
        // Write via the same source-gen path that TypedExport uses.
        await TypedExport.Run(outputPath, records);

        // Read it back via the source-gen ReadRows extension.
        using var wb = await Workbook.OpenAsync(outputPath);
        var sheet = wb[SheetName];
        return sheet.ReadRows().ToList();
    }
}
