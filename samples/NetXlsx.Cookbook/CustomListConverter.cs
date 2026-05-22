// Cookbook recipe — CustomListConverter (v1.1 / decision I-58).
//
// A tag-list property (List<string>) gets serialized as a single
// semicolon-delimited cell via a user-supplied ICellConverter<T>.
// Demonstrates the v1.1 typed-mapping escape hatch for property
// types the generator's built-in scalar set doesn't cover.

using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a workbook with a "Issues" sheet where each row carries a
/// list of tags in column C. The list serializes as
/// <c>"tag-a;tag-b;tag-c"</c> via the user-defined
/// <see cref="TagListConverter"/>.
/// </summary>
public static class CustomListConverter
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Issues";

    /// <summary>Encodes / decodes a <c>List&lt;string&gt;</c> as a semicolon-delimited cell.</summary>
    public sealed class TagListConverter : ICellConverter<List<string>>
    {
        public void Write(ICell cell, List<string> value) =>
            cell.SetString(value is null ? "" : string.Join(";", value));

        public List<string> Read(ICell cell)
        {
            var s = cell.GetString();
            return string.IsNullOrEmpty(s) ? new List<string>() : s.Split(';').ToList();
        }
    }

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet(SheetName);

        // Header row + a couple of data rows written via raw IRow.Set for
        // clarity. The custom-converter path runs when callers use the
        // source-generated AddRow / ReadRows extensions on Issue. The
        // recipe leaves that pattern to the docs since the converter is
        // the surface being demoed; this recipe shows the cell layout.
        sh.AppendRow().Set(1, "Id").Set(2, "Title").Set(3, "Tags");

        var conv = new TagListConverter();
        var row1 = sh.AppendRow().Set(1, 101).Set(2, "Fuzz harness finding — IOoR leak");
        conv.Write(row1.Cell(3), new List<string> { "security", "open-path", "v1.1" });

        var row2 = sh.AppendRow().Set(1, 102).Set(2, "Bench breadth landed");
        conv.Write(row2.Cell(3), new List<string> { "perf", "bench" });

        sh.FreezeRows(1);
        await wb.SaveAsync(outputPath);
    }
}
