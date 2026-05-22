// Cookbook recipe — RichTextCells (v1.1 / decision I-50).
//
// Multi-run formatted strings in a single cell. Per-run typography
// comes from RichTextStyle (font-only subset of CellStyle); cell-level
// fills/borders still go through ICell.Style(CellStyle).

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a one-sheet "release announcement" workbook where each cell
/// in the body mixes multiple text styles inline — bold callouts, an
/// italic clarifier, a colored emphasis run.
/// </summary>
public static class RichTextCells
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Announcement";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet(SheetName);

        // Single cell, three runs: bold "Release:", normal "1.1.0", italic note.
        sh["A1"].SetRichText(new RichText(
            new RichTextRun("Release: ", new RichTextStyle { Bold = true }),
            new RichTextRun("1.1.0"),
            new RichTextRun("  — features only; tag pending review.", new RichTextStyle { Italic = true })));

        // Single cell, mixed colors + sizes for emphasis.
        sh["A2"].SetRichText(new RichText(
            new RichTextRun("Status: ", new RichTextStyle { Bold = true }),
            new RichTextRun("GREEN", new RichTextStyle { Bold = true, Color = Color.Green, FontSize = 14 })));

        // Single cell, underlined link-like run.
        sh["A3"].SetRichText(new RichText(
            new RichTextRun("See: ", new RichTextStyle { Italic = true }),
            new RichTextRun("docs/design.md §6.8.1",
                new RichTextStyle { Underline = UnderlineStyle.Single, Color = Color.Blue })));

        await wb.SaveAsync(outputPath);
    }
}
