// Cookbook recipe — MacroEnabledPassthrough (post-v1.1 / decision I-69).
//
// NetXlsx is .xlsx-focused, but many real workbooks carry VBA macros
// (.xlsm). Workbook.CreateMacroEnabled() produces a macro-enabled
// workbook; Workbook.Open transparently reads either format. Any VBA
// project present in an opened .xlsm is preserved byte-for-byte across a
// save — NetXlsx never reads, writes, or executes VBA, it only passes it
// through. IsMacroEnabled reports the format. Save such workbooks with an
// .xlsm extension (this recipe is given one by its caller).

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a macro-enabled (.xlsm) workbook with a single data sheet.
/// There is no public API to author a VBA project — the value of the
/// macro-enabled format here is round-trip preservation of an existing
/// one — so the recipe simply demonstrates the create / IsMacroEnabled
/// path that the passthrough guarantee rests on.
/// </summary>
public static class MacroEnabledPassthrough
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Data";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.CreateMacroEnabled();

        // IsMacroEnabled is true from creation; it would also be true for
        // any .xlsm opened via Workbook.Open, with its VBA project intact.
        var sh = wb.AddSheet(SheetName);
        sh.AppendRow().Set(1, "Metric").Set(2, "Value");
        sh.AppendRow().Set(1, "MacroEnabled").Set(2, wb.IsMacroEnabled ? "yes" : "no");

        await wb.SaveAsync(outputPath);
    }
}
