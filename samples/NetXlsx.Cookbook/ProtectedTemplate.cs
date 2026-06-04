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
using S = DocumentFormat.OpenXml.Spreadsheet;

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
        // reach through .Underlying since the wrapper doesn't expose a
        // Locked flag on CellStyle. (A candidate public surface addition;
        // queued post-v2.0.0.) In OOXML that is a cellXfs <xf> carrying
        // <protection locked="0"/> — append one and point the cells at it.
        var cellXfs = wb.Underlying.WorkbookPart!.WorkbookStylesPart!
            .Stylesheet!.GetFirstChild<S.CellFormats>()!;
        cellXfs.AppendChild(new S.CellFormat
        {
            NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0,
            ApplyProtection = true,
            Protection = new S.Protection { Locked = false },
        });
        cellXfs.Count = (uint)cellXfs.ChildElements.Count;
        uint unlockedXf = cellXfs.Count.Value - 1;
        for (int r = 2; r <= 4; r++)
            inputs[r, 2].Underlying.StyleIndex = unlockedXf;
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
