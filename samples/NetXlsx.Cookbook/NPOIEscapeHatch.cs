// Cookbook recipe 10 — NPOIEscapeHatch
//
// Per docs/design.md §8.1: "Use .Underlying to do something the facade
// doesn't cover (e.g., set a print area)."
//
// Demonstrates the design's first-class-escape-hatch promise
// (decisions #1, #32): every wrapper type exposes its raw NPOI
// counterpart, so the facade is additive over NPOI, not a sandbox.
// This recipe sets a print area — something v1 deliberately doesn't
// model — through the escape hatch.

using System.Threading.Tasks;
using NetXlsx;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a small data sheet, then reaches through
/// <see cref="ISheet.Underlying"/> to set a print area and configure
/// landscape page orientation — operations NetXlsx does not wrap
/// in v1. The wrapper still owns the workbook lifecycle; the escape
/// hatch is for incremental capability, not workaround.
/// </summary>
public static class NPOIEscapeHatch
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "PrintMe";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet(SheetName);

        sheet.AppendRow().Set(1, "Region").Set(2, "Q1").Set(3, "Q2").Set(4, "Q3").Set(5, "Q4");
        sheet.AppendRow().Set(1, "North").Set(2, 1200m).Set(3, 1500m).Set(4, 1700m).Set(5, 1600m);
        sheet.AppendRow().Set(1, "South").Set(2, 1100m).Set(3, 1300m).Set(4, 1450m).Set(5, 1550m);
        sheet.AppendRow().Set(1, "East").Set(2, 1400m).Set(3, 1600m).Set(4, 1800m).Set(5, 1900m);
        sheet.AppendRow().Set(1, "West").Set(2, 1050m).Set(3, 1250m).Set(4, 1400m).Set(5, 1500m);

        // ---- Escape hatch territory ----------------------------------
        var rawWorkbook = wb.Underlying;
        var rawSheet = sheet.Underlying;

        // 1. Print area = A1:E5 — every cell we wrote.
        //    Workbook-level SetPrintArea takes a 0-based sheet index
        //    and column/row bounds (also 0-based, inclusive).
        int sheetIndex = rawWorkbook.GetSheetIndex(rawSheet);
        rawWorkbook.SetPrintArea(
            sheetIndex,
            startColumn: 0, endColumn: 4,
            startRow:    0, endRow:    4);

        // 2. Landscape orientation and "fit to 1 page wide".
        var printSetup = rawSheet.PrintSetup;
        printSetup.Landscape = true;
        printSetup.FitWidth  = 1;
        printSetup.FitHeight = 0;            // 0 == unbounded; print as many pages tall as needed
        rawSheet.FitToPage = true;

        // 3. Header/footer.
        rawSheet.Header.Center = "Regional sales — annual";
        rawSheet.Footer.Right  = "Page &P of &N";

        // 4. Repeat the header row on every printed page (NPOI's
        //    RepeatingRows API — also unmodeled in v1).
        rawSheet.RepeatingRows = new CellRangeAddress(0, 0, -1, -1);

        await wb.SaveAsync(outputPath);
    }
}
