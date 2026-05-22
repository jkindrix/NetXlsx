// Cookbook recipe — ProtectedTemplate (v1.1 / decisions I-53, I-54).
//
// A common shape for shared spreadsheet templates: most cells are
// locked so the user can only edit a designated input column. Combines
// sheet-level protection (locks cell editing) with workbook-level
// structure protection (prevents sheet add/delete/rename).
//
// Excel's sheet protection is a UX guard, not security — the password
// hash is weak. The recipe documents the trade-off in its comments
// so readers don't take "password" for cryptographic protection.

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a template workbook with a "Reference" sheet (fully locked)
/// and an "Inputs" sheet (locked except for column B). Workbook
/// structure is locked so users can't add or remove sheets.
/// </summary>
public static class ProtectedTemplate
{
    /// <summary>Output sheet name (the inputs sheet).</summary>
    public const string SheetName = "Inputs";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();

        var refSheet = wb.AddSheet("Reference");
        refSheet.AppendRow().Set(1, "Constant").Set(2, "Value");
        refSheet.AppendRow().Set(1, "PI").Set(2, 3.14159);
        refSheet.AppendRow().Set(1, "E").Set(2, 2.71828);
        // Fully lock the reference sheet — all defaults restrictive.
        refSheet.Protect(options: SheetProtection.LockAll);

        var inputs = wb.AddSheet(SheetName);
        inputs.AppendRow().Set(1, "Label").Set(2, "Value");
        inputs.AppendRow().Set(1, "Customer name");
        inputs.AppendRow().Set(1, "Customer ID");
        inputs.AppendRow().Set(1, "Order amount");

        // For input cells, mark them unlocked at the cell-style level —
        // reach through .Underlying since the wrapper doesn't yet expose
        // a Locked flag on CellStyle. (Likely a v1.2 surface addition.)
        for (int r = 2; r <= 4; r++)
        {
            var c = inputs[r, 2];
            var cs = c.Underlying.Sheet.Workbook.CreateCellStyle();
            cs.IsLocked = false;
            c.Underlying.CellStyle = cs;
        }
        // Allow user to type into unlocked cells; block format / structure.
        inputs.Protect(password: "user-template", options: new SheetProtection
        {
            LockFormatCells = true,
            LockInsertRows = true,
            LockDeleteRows = true,
        });

        // Workbook-level: prevent adding / removing / renaming sheets.
        wb.Protect(WorkbookProtection.LockStructure);

        await wb.SaveAsync(outputPath);
    }
}
