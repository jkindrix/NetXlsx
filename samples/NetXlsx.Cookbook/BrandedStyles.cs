// Cookbook recipe — BrandedStyles (v1.1 / decision I-57).
//
// Register a small palette of named CellStyle values on the workbook,
// then apply them by name on cells and ranges. The intended use case
// is "company branding" — a fixed set of header / accent / footer
// styles reused across many reports.
//
// Reminder: v1.1 named styles are an in-process convenience. The
// styles themselves still serialize via the style-pool dedup, but
// the name -> style map is not rehydrated by Workbook.Open.

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a one-sheet quarterly-report workbook where every cell uses
/// one of three named styles ("BrandHeader", "BrandBody",
/// "BrandFooter") — all registered at the top of Run.
/// </summary>
public static class BrandedStyles
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Q2Report";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();

        wb.RegisterStyle("BrandHeader", new CellStyle
        {
            Bold = true,
            FontSize = 14,
            FontColor = Color.White,
            Background = Color.FromRgb(0x2C, 0x3E, 0x50),
            HorizontalAlignment = HAlign.Center,
        });
        wb.RegisterStyle("BrandBody", new CellStyle
        {
            FontSize = 11,
            Background = Color.FromRgb(0xF4, 0xF6, 0xF8),
        });
        wb.RegisterStyle("BrandFooter", new CellStyle
        {
            Italic = true,
            FontSize = 9,
            FontColor = Color.Gray,
        });

        var sh = wb.AddSheet(SheetName);
        sh.AppendRow().Set(1, "Region").Set(2, "Revenue");
        sh.Range("A1:B1").ApplyNamedStyle("BrandHeader");

        sh.AppendRow().Set(1, "EU").Set(2, 125_000);
        sh.AppendRow().Set(1, "US").Set(2, 215_000);
        sh.Range("A2:B3").ApplyNamedStyle("BrandBody");

        sh.AppendRow().Set(1, "Source: internal CRM, 2026-Q2");
        sh["A4"].ApplyNamedStyle("BrandFooter");

        await wb.SaveAsync(outputPath);
    }
}
