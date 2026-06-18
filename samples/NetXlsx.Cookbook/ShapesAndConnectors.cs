// Cookbook recipe — ShapesAndConnectors (post-v1.1 / decisions I-74, I-80).
//
// Vector drawing primitives. AddShape draws a preset geometry (rectangle,
// ellipse, …) anchored between two cells with optional fill/line colors.
// AddConnector draws a connector line — straight / bent / curved — with
// optional arrowheads (ConnectorEnd), flip flags, and a line width in
// points. Together they sketch a tiny flowchart.

using System.Threading.Tasks;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a one-sheet two-box flowchart: a rounded "Start" rectangle and
/// an "End" ellipse, joined by a straight arrow connector. Shows both
/// the shape and connector surfaces in one diagram.
/// </summary>
public static class ShapesAndConnectors
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "Flowchart";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet(SheetName);

        sh["B2"].SetString("Start");
        sh["B8"].SetString("End");

        // "Start" — a filled rounded rectangle over B2:D4.
        sh.AddShape(
            ShapeType.RoundedRectangle, "B2", "D4",
            fillColor: Color.FromRgb(0xDA, 0xE8, 0xFC),
            lineColor: Color.FromRgb(0x6C, 0x8E, 0xBF));

        // "End" — an ellipse over B8:D10.
        sh.AddShape(
            ShapeType.Ellipse, "B8", "D10",
            fillColor: Color.FromRgb(0xD5, 0xE8, 0xD4),
            lineColor: Color.FromRgb(0x82, 0xB3, 0x66));

        // A straight arrow connecting the bottom of "Start" to "End",
        // 2 pt wide with an arrowhead at the tail.
        sh.AddConnector(
            ConnectorType.Straight, "C4", "C8",
            lineColor: Color.FromRgb(0x66, 0x66, 0x66),
            tailEnd: ConnectorEnd.Arrow,
            lineWidthPoints: 2.0);

        await wb.SaveAsync(outputPath);
    }
}
