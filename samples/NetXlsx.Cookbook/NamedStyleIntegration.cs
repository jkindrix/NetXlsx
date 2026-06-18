// Cookbook recipe — NamedStyleIntegration (post-v1.1 / decision I-67).
//
// Named styles that round-trip through the OOXML file format. You
// register a CellStyle under a name and apply it by name (as in the
// BrandedStyles recipe). What I-67 adds over the v1.1 in-process
// convenience (I-57) is real OOXML integration: RegisterStyle now writes
// a <cellStyle> entry into styles.xml (the table behind Excel's "Cell
// Styles" ribbon group), so Workbook.Open rehydrates the name -> style
// map. Reopen a saved file and RegisteredStyleNames still lists the
// names; GetRegisteredStyle returns the style. Cells styled via
// ApplyNamedStyle dedup to one shared cellXfs entry on disk.

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a one-sheet report that registers a "Heading" named style,
/// applies it to two separate cells, and saves. The style name survives
/// a round-trip (it lands in the OOXML cellStyles table), and the two
/// headings collapse to a single shared style-pool entry in the file.
/// </summary>
public static class NamedStyleIntegration
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Report";

    /// <summary>The registered style name.</summary>
    public const string HeadingStyle = "Heading";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();

        wb.RegisterStyle(HeadingStyle, new CellStyle
        {
            Bold = true,
            FontSize = 13,
            FontColor = Color.White,
            Background = Color.FromRgb(0x34, 0x49, 0x55),
        });

        var sh = wb.AddSheet(SheetName);

        // Two section headings reuse the one registered style. On disk
        // these resolve to a single shared cellXfs entry, and the name
        // "Heading" is written to the cellStyles table.
        sh["A1"].SetString("Summary");
        sh["A1"].ApplyNamedStyle(HeadingStyle);
        sh.AppendRow();
        sh.AppendRow().Set(1, "Detail");
        sh["A3"].ApplyNamedStyle(HeadingStyle);

        sh["A4"].SetString("(reopening this file keeps the bold heading look");
        sh["A5"].SetString(" AND RegisteredStyleNames still lists \"Heading\" — I-67)");

        await wb.SaveAsync(outputPath);
    }
}
