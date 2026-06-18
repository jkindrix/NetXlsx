// Cookbook recipe — ConditionalFormatting (post-v1.1 / decision I-73).
//
// Highlight cells by rule. NetXlsx covers the cell-value comparison
// rules (greater-than, between, …), an arbitrary-formula rule, and
// 2-/3-color scales. This recipe shows one of each so the three
// rule families are visible side-by-side; the ConditionalFormat
// factories are the only way to build a rule (no public ctor).

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a one-sheet score report whose grades column carries three
/// kinds of conditional formatting: a cell-value rule (fail &lt; 50),
/// a formula rule (flag the row when a "Retake" marker is set), and a
/// 3-color scale across the whole grades column.
/// </summary>
public static class ConditionalFormatting
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Scores";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet(SheetName);

        sh.AppendRow().Set(1, "Student").Set(2, "Score").Set(3, "Retake");
        sh.AppendRow().Set(1, "Ada").Set(2, 92).Set(3, "");
        sh.AppendRow().Set(1, "Bert").Set(2, 41).Set(3, "Y");
        sh.AppendRow().Set(1, "Cleo").Set(2, 68).Set(3, "");
        sh.AppendRow().Set(1, "Dane").Set(2, 55).Set(3, "");

        // Rule 1 — cell-value: a failing score (< 50) goes bold red.
        sh.AddConditionalFormatting(
            "B2:B5",
            ConditionalFormat.CellValueLessThan(
                "50",
                new CellStyle { Bold = true, FontColor = Color.FromRgb(0xC0, 0x00, 0x00) }));

        // Rule 2 — formula: shade the score when the same row's Retake
        // marker (column C) is non-empty. $C2 is relative-row, fixed-col.
        sh.AddConditionalFormatting(
            "B2:B5",
            ConditionalFormat.Formula(
                "$C2<>\"\"",
                new CellStyle { Background = Color.FromRgb(0xFF, 0xF2, 0xCC) }));

        // Rule 3 — 3-color scale: red (low) → yellow (mid) → green (high).
        sh.AddConditionalFormatting(
            "B2:B5",
            ConditionalFormat.ColorScale(
                Color.FromRgb(0xF8, 0x69, 0x6B),
                Color.FromRgb(0xFF, 0xEB, 0x84),
                Color.FromRgb(0x63, 0xBE, 0x7B)));

        await wb.SaveAsync(outputPath);
    }
}
