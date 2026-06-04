// Cookbook recipe 13 — CellErrors
//
// Per docs/design.md §8.1: "Read a workbook containing formula errors;
// classify them via GetError() and the CellError enum."
//
// Writes a workbook with each of the eight standard Excel error literals
// (via the .Underlying escape hatch — no SetError API) and reads them back
// classified. OOXML stores an error cell as t="e" with the literal in <v>;
// since v2.0.0 the hatch is the SDK <c> element, so the recipe authors
// exactly that shape — including #GETTING_DATA, which the old NPOI write
// API could not produce.

using System.Threading.Tasks;
using NetXlsx;
using S = DocumentFormat.OpenXml.Spreadsheet;

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

        foreach (var (literal, error, description) in Errors)
        {
            var row = sheet.AppendRow();
            // Write the error via .Underlying — no SetError on ICell.
            // An OOXML error cell is t="e" with the literal in <v>.
            var c = row.Cell(1).Underlying;
            c.DataType = S.CellValues.Error;
            c.CellValue = new S.CellValue(literal);
            row.Set(2, description);
            _ = error;   // present in the tuple for clarity; not used in the write.
        }

        await wb.SaveAsync(outputPath);
    }

    /// <summary>
    /// All eight Excel error literals, including <c>#GETTING_DATA</c> —
    /// authorable through the SDK escape hatch since v2.0.0 (the legacy
    /// NPOI write API refused it; Excel produces it from external data
    /// sources).
    /// </summary>
    public static readonly (string Literal, CellError EnumValue, string Description)[] Errors =
    {
        ("#NULL!", CellError.Null,                "#NULL! — intersection of two ranges that do not intersect"),
        ("#DIV/0!", CellError.DivByZero,          "#DIV/0! — division by zero"),
        ("#VALUE!", CellError.Value,              "#VALUE! — wrong type of operand"),
        ("#REF!", CellError.Ref,                  "#REF! — reference is not valid"),
        ("#NAME?", CellError.Name,                "#NAME? — unrecognized name in a formula"),
        ("#NUM!", CellError.Num,                  "#NUM! — invalid numeric value"),
        ("#N/A", CellError.NotAvailable,          "#N/A — value is not available"),
        ("#GETTING_DATA", CellError.GettingData,  "#GETTING_DATA — data fetching from an external source"),
    };
}
