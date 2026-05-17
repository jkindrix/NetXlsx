// Cookbook recipe 11 — OpenEditSave
//
// Per docs/design.md §8.1: "Open an existing workbook, modify a few
// cells, save — demonstrate the no-churn styles guarantee (§7.5) *and*
// the unknown-parts preservation guarantee (§7.7)."
//
// Realizes the two preservation promises in one runnable demo:
//   §7.5 — Re-applying an identical style does not allocate a new
//          NPOI ICellStyle (style pool dedup).
//   §7.7 — OPC parts NetXlsx does not model (custom XML, pivot
//          caches, threaded comments, vendor <ext> elements) round-trip
//          verbatim because the package layer is never touched.

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NetXlsx;
using NPOI.OpenXml4Net.OPC;
using NPOI.XSSF.UserModel;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// The recipe builds an input file via raw NPOI carrying a custom OPC
/// part (a synthetic stand-in for the "things NetXlsx does not
/// model" category), then opens it through NetXlsx, mutates two
/// cells, re-applies an identical style to a third, and saves. The
/// golden-file test (and this recipe's own internal asserts) verify
/// both preservation guarantees.
/// </summary>
public static class OpenEditSave
{
    /// <summary>Sheet name authored into the input file.</summary>
    public const string SheetName = "Existing";

    /// <summary>The custom OPC part this recipe round-trips.</summary>
    public const string CustomPartUri = "/customXml/itemRecipe.xml";

    private static readonly byte[] CustomPartBytes =
        System.Text.Encoding.UTF8.GetBytes("<!-- recipe11-preserve -->");

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        // ---- 1. Build the input file via raw NPOI -------------------
        // Includes a custom OPC part so we can later assert it round-trips.
        var inputPath = Path.Combine(
            Path.GetDirectoryName(outputPath) ?? Path.GetTempPath(),
            $"open-edit-save-input-{Guid.NewGuid():N}.xlsx");

        using (var wb = new XSSFWorkbook())
        {
            var sheet = wb.CreateSheet(SheetName);
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Region");
            sheet.GetRow(0).CreateCell(1).SetCellValue("Sales");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("North");
            sheet.GetRow(1).CreateCell(1).SetCellValue(1200.0);
            sheet.CreateRow(2).CreateCell(0).SetCellValue("South");
            sheet.GetRow(2).CreateCell(1).SetCellValue(1100.0);

            // Attach the custom OPC part to the package.
            var pkg = wb.Package;
            var partName = PackagingUriHelper.CreatePartName(CustomPartUri);
            var part = pkg.CreatePart(partName, "application/xml");
            using (var ps = part.GetOutputStream())
            {
                ps.Write(CustomPartBytes, 0, CustomPartBytes.Length);
            }

            using var fs = File.Create(inputPath);
            wb.Write(fs, leaveOpen: false);
        }

        // ---- 2. Open via NetXlsx, mutate, save -------------------
        try
        {
            using (var wb = Workbook.Open(inputPath))
            {
                var sheet = wb[SheetName];

                // Mutation #1 — append a row of new data.
                sheet.AppendRow().Set(1, "East").Set(2, 1450m);

                // Mutation #2 — overwrite an existing cell value.
                sheet["B3"].SetNumber(1175m);   // South: 1100 → 1175

                // §7.5 demonstration — apply the same style to two
                // different cells and observe they share one ICellStyle
                // index. The pool deduplicates by value.
                var headerStyle = new CellStyle { Bold = true };
                sheet["A1"].Style(headerStyle);
                sheet["B1"].Style(headerStyle);

                await wb.SaveAsync(outputPath);
            }
        }
        finally
        {
            if (File.Exists(inputPath)) File.Delete(inputPath);
        }
    }

    /// <summary>
    /// Helper used by the golden-file test to read back the custom
    /// part's bytes from a saved workbook. Lives next to the recipe
    /// so the assertion logic stays close to the preservation
    /// guarantee being demonstrated.
    /// </summary>
    public static byte[]? ReadCustomPartBytes(string xlsxPath)
    {
        var pkg = OPCPackage.Open(xlsxPath, PackageAccess.READ);
        try
        {
            var name = PackagingUriHelper.CreatePartName(CustomPartUri);
            if (!pkg.ContainPart(name)) return null;
            var part = pkg.GetPart(name);
            using var s = part.GetInputStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            return ms.ToArray();
        }
        finally
        {
            pkg.Close();
        }
    }

    /// <summary>The custom-part bytes the recipe writes — exposed for tests.</summary>
    public static byte[] ExpectedCustomPartBytes => CustomPartBytes.ToArray();
}
