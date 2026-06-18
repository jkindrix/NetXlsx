// Cookbook recipe — PictureBorders (post-v1.1 / decision I-86).
//
// An embedded picture can carry a border (the OOXML <a:ln> on the
// picture's shape properties). IPicture.Border takes a PictureBorder
// record: an explicit sRGB Color (or a theme slot), plus an optional
// WidthPoints. Setting Border = null removes any border. The border
// round-trips: reopening the workbook rehydrates IPicture.Border.

using System;
using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a one-sheet workbook with a single embedded PNG that carries a
/// 1.5 pt navy border. Uses the canonical tiny 1×1 PNG as the byte
/// source so no external fixture file is needed.
/// </summary>
public static class PictureBorders
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Framed";

    /// <summary>The border width applied to the picture, in points.</summary>
    public const double BorderWidthPoints = 1.5;

    // 67-byte canonical 1×1 transparent PNG.
    private static readonly byte[] s_tinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk" +
        "YAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet(SheetName);

        sh["A1"].SetString("Bordered logo:");
        var pic = sh.AddPicture("B1", s_tinyPng);

        // Frame the picture with a 1.5 pt navy border.
        pic.Border = new PictureBorder
        {
            Color = Color.FromRgb(0x1F, 0x38, 0x64),
            WidthPoints = BorderWidthPoints,
        };

        await wb.SaveAsync(outputPath);
    }
}
