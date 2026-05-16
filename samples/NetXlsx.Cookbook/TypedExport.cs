// Cookbook recipe 3 — TypedExport
//
// Per docs/design.md §8.1: "Same as [TabularExport] using [Worksheet] +
// source-gen extension methods."
//
// v0.3.x: demonstrates the source-generated typed-mapping write path.
// The [Worksheet] type below produces the source-gen extension class
// `SalesRecord_SheetExtensions` with AddRow/AddRows methods, which
// internally call ISheet.AppendRow + IRow.Set. Compare with
// TabularExport.cs to see the typed-mapping ergonomics:
//   - No column index arithmetic in the recipe code.
//   - Property-to-column ordering follows source declaration order
//     (v0.3.x — [Column(Order)] is parsed but reorder-on-write lands
//     with the styling slice).
//   - The header row is still written by hand; auto-header on AddRows
//     is a follow-up enhancement.

using System.Collections.Generic;
using System.Threading.Tasks;
using NetXlsx;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// A typed worksheet row schema. The generator emits a
/// <c>SalesRecord_SheetExtensions</c> static class with
/// <c>AddRow(this ISheet, SalesRecord)</c> and
/// <c>AddRows(this ISheet, IEnumerable&lt;SalesRecord&gt;)</c>.
/// </summary>
[Worksheet(Visibility = WorksheetVisibility.Public)]
public partial record SalesRecord
{
    /// <summary>Sales region.</summary>
    [Column("Region", Order = 0)] public string Region { get; init; } = "";
    /// <summary>Total revenue.</summary>
    [Column("Revenue", Order = 1, Format = "$#,##0.00")] public decimal Revenue { get; init; }
    /// <summary>Margin (0..1).</summary>
    [Column("Margin", Order = 2)] public double Margin { get; init; }
    /// <summary>True when this row is flagged strategic.</summary>
    [Column("Strategic", Order = 3)] public bool Strategic { get; init; }
}

/// <summary>
/// Cookbook recipe writing a typed dataset via the [Worksheet] source
/// generator's emitted extension methods.
/// </summary>
public static class TypedExport
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Typed Sales";

    /// <summary>Runs the recipe with a small fixed dataset.</summary>
    public static Task Run(string outputPath) =>
        Run(outputPath, new[]
        {
            new SalesRecord { Region = "North", Revenue = 1000.50m, Margin = 0.12, Strategic = true  },
            new SalesRecord { Region = "South", Revenue = 2500.00m, Margin = 0.18, Strategic = false },
            new SalesRecord { Region = "East",  Revenue = 3700.75m, Margin = 0.22, Strategic = true  },
        });

    /// <summary>Runs the recipe with a caller-supplied dataset.</summary>
    public static async Task Run(string outputPath, IReadOnlyList<SalesRecord> records)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet(SheetName);

        // Header — still authored by hand in v0.3.x; auto-header on
        // AddRows is a follow-up enhancement.
        sheet.AppendRow()
            .Set(1, "Region")
            .Set(2, "Revenue")
            .Set(3, "Margin")
            .Set(4, "Strategic");

        // Typed write — single call per record, source-generated.
        sheet.AddRows(records);

        await wb.SaveAsync(outputPath);
    }
}
