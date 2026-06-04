// Cookbook recipe 10 — OpenXmlEscapeHatch
//
// Per docs/design.md §8.1: "Use .Underlying to do something the facade
// doesn't cover (e.g., set a print area)."
//
// Demonstrates the design's first-class-escape-hatch promise
// (decisions #1, #32 / I-82): every wrapper type exposes its raw Open XML
// SDK counterpart, so the facade is additive over the OOXML document, not
// a sandbox. This recipe sets a print area, page setup, header/footer and
// repeating title rows — none of which v2 deliberately models — straight
// onto the SDK DOM.
//
// (Until the v2.0.0 engine cutover this recipe was NPOIEscapeHatch and
// drove the same artifacts through NPOI's API; the hatch retyped to the
// SDK at the cutover, decision I-82.)

using System.Threading.Tasks;
using NetXlsx;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a small data sheet, then reaches through
/// <see cref="IWorkbook.Underlying"/> / <see cref="ISheet.Underlying"/> to
/// set a print area and configure landscape page orientation — operations
/// NetXlsx does not wrap. The wrapper still owns the workbook lifecycle;
/// the escape hatch is for incremental capability, not workaround.
/// </summary>
public static class OpenXmlEscapeHatch
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
        var workbook = wb.Underlying.WorkbookPart!.Workbook!;
        var worksheet = sheet.Underlying;

        // 1. Print area + repeating title row. Both are the built-in defined
        //    names Excel itself writes (_xlnm.Print_Area / _xlnm.Print_Titles),
        //    scoped to the sheet via localSheetId. <definedNames> must sit
        //    after <sheets> in CT_Workbook's child sequence.
        var definedNames = new S.DefinedNames(
            new S.DefinedName($"{SheetName}!$A$1:$E$5") { Name = "_xlnm.Print_Area", LocalSheetId = 0 },
            new S.DefinedName($"{SheetName}!$1:$1") { Name = "_xlnm.Print_Titles", LocalSheetId = 0 });
        workbook.InsertAfter(definedNames, workbook.GetFirstChild<S.Sheets>());

        // 2. Landscape orientation and "fit to 1 page wide". fitToPage lives
        //    in <sheetPr> (the FIRST CT_Worksheet child); <pageSetup> goes at
        //    the tail, before <headerFooter>.
        worksheet.InsertAt(new S.SheetProperties(new S.PageSetupProperties { FitToPage = true }), 0);
        worksheet.AppendChild(new S.PageSetup
        {
            Orientation = S.OrientationValues.Landscape,
            FitToWidth = 1,
            FitToHeight = 0, // 0 == unbounded; print as many pages tall as needed
        });

        // 3. Header/footer — Excel's &-token syntax, verbatim.
        worksheet.AppendChild(new S.HeaderFooter
        {
            OddHeader = new S.OddHeader("&CRegional sales — annual"),
            OddFooter = new S.OddFooter("&RPage &P of &N"),
        });

        await wb.SaveAsync(outputPath);
    }
}
