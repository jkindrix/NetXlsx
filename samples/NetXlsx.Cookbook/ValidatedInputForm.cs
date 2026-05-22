// Cookbook recipe — ValidatedInputForm (v1.1 / decision I-55).
//
// Builds a form-like sheet with one validated cell per input type:
// a dropdown (list), a numeric range, a date range, a text-length
// cap, and a custom-formula rule. Each demonstrates one of the
// DataValidation static factories.

using System;
using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a single-sheet "submission form" workbook where each
/// input cell has a different kind of validation. Open in Excel
/// to see the rules in action — invalid input triggers the default
/// error dialog.
/// </summary>
public static class ValidatedInputForm
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "SubmissionForm";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet(SheetName);

        sh.AppendRow().Set(1, "Field").Set(2, "Value");

        // Row 2: list / dropdown
        sh.AppendRow().Set(1, "Status");
        sh.AddValidation("B2", DataValidation.List("Open", "In Progress", "Closed"));

        // Row 3: integer between 1 and 100
        sh.AppendRow().Set(1, "Priority (1-100)");
        sh.AddValidation("B3", DataValidation.IntegerBetween(1, 100));

        // Row 4: date in current year (illustrative; uses a fixed window
        // here so the recipe output is deterministic for golden-file).
        sh.AppendRow().Set(1, "Submitted (2026)");
        sh.AddValidation("B4", DataValidation.DateBetween(
            new DateOnly(2026, 1, 1), new DateOnly(2026, 12, 31)));

        // Row 5: text-length cap (subject line under 80 chars)
        sh.AppendRow().Set(1, "Subject (<=80 chars)");
        sh.AddValidation("B5", DataValidation.TextLengthAtMost(80));

        // Row 6: custom — value must be a number (any).
        sh.AppendRow().Set(1, "Numeric (any number)");
        sh.AddValidation("B6", DataValidation.Custom("ISNUMBER(B6)"));

        sh.FreezeRows(1);

        await wb.SaveAsync(outputPath);
    }
}
