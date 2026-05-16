// Cookbook recipe 13 — CellErrors
//
// Per docs/design.md §8.1: "Read a workbook containing formula errors;
// classify them via GetError() and the CellError enum."
//
// v0.4 cell-errors slice: writes a workbook with each of the eight
// standard Excel error codes (via the .Underlying escape hatch — no
// SetError API yet) and reads them back classified.

using System.Threading.Tasks;
using NetXlsx;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Writes one row per Excel error code, with the error in column A and
/// a human-readable description in column B. The recipe's value is at
/// read time — call <see cref="ICell.GetError"/> on column A and you
/// get a typed <see cref="NetXlsx.CellError"/> enum value rather
/// than a stringly-typed <c>"#DIV/0!"</c>.
/// </summary>
public static class CellErrors
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Errors";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet(SheetName);

        sheet.AppendRow()
            .Set(1, "Error code")
            .Set(2, "Description");

        foreach (var (code, error, description) in Errors)
        {
            var row = sheet.AppendRow();
            // Write the error via .Underlying — no SetError on ICell yet.
            row.Cell(1).Underlying.SetCellErrorValue(code);
            row.Set(2, description);
            _ = error;   // present in the tuple for clarity; not used in the write.
        }

        await wb.SaveAsync(outputPath);
    }

    /// <summary>
    /// The seven Excel error codes NPOI can write programmatically.
    /// <c>#GETTING_DATA</c> (0x2B) is unreachable from NPOI's
    /// <c>SetCellErrorValue</c> — only Excel itself can produce it from
    /// external data sources — but <see cref="CellError.GettingData"/>
    /// is still surfaced on the read side for workbooks authored by Excel.
    /// </summary>
    public static readonly (byte ExcelCode, CellError EnumValue, string Description)[] Errors =
    {
        (0x00, CellError.Null,         "#NULL! — intersection of two ranges that do not intersect"),
        (0x07, CellError.DivByZero,    "#DIV/0! — division by zero"),
        (0x0F, CellError.Value,        "#VALUE! — wrong type of operand"),
        (0x17, CellError.Ref,          "#REF! — reference is not valid"),
        (0x1D, CellError.Name,         "#NAME? — unrecognized name in a formula"),
        (0x24, CellError.Num,          "#NUM! — invalid numeric value"),
        (0x2A, CellError.NotAvailable, "#N/A — value is not available"),
    };
}
