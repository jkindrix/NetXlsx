using System;
using System.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.PublicApi;

/// <summary>
/// Companion assertions to the <c>Microsoft.CodeAnalysis.PublicApiAnalyzers</c>
/// wiring on the main library. The analyzer catches additions at compile
/// time via <c>RS0016</c>/<c>RS0017</c> diagnostics (treated as errors —
/// see <c>.editorconfig</c>). These tests provide a runtime backstop:
/// a guarded enumeration of the public surface that fails if any public
/// type leaks before being deliberately added.
/// </summary>
public class PublicApiSnapshotTests
{
    [Fact]
    public void Public_Surface_Matches_Expected_Baseline()
    {
        // Baseline: the source-generator marker attributes and the empty
        // ISheet/IWorkbook stubs. Additional public types land alongside
        // entries in PublicAPI.Unshipped.txt. This test catches accidental
        // leaks at runtime; the analyzer (RS0016/RS0017) catches them at
        // compile time. Both are deliberate.
        var assembly = System.Reflection.Assembly.Load("NetXlsx");
        var publicTypeNames = assembly
            .GetExportedTypes()
            .Where(t => !IsCompilerGenerated(t))
            .Select(t => t.FullName)
            .OrderBy(n => n, System.StringComparer.Ordinal)
            .ToArray();

        var expected = new[]
        {
            "NetXlsx.BorderStyle",
            "NetXlsx.CellAddress",
            "NetXlsx.CellBorders",
            "NetXlsx.CellError",
            "NetXlsx.CellKind",
            "NetXlsx.CellStyle",
            "NetXlsx.ChartType",
            "NetXlsx.Color",
            "NetXlsx.ColumnAttribute",
            "NetXlsx.ConditionalFormat",
            "NetXlsx.ConnectorEnd",
            "NetXlsx.ConnectorType",
            "NetXlsx.DataValidation",
            "NetXlsx.DateSystem",
            "NetXlsx.FilterCriteria",
            "NetXlsx.FormulaException",
            "NetXlsx.HAlign",
            "NetXlsx.ICell",
            "NetXlsx.ICellConverter`1",
            "NetXlsx.IChart",
            "NetXlsx.IColumn",
            "NetXlsx.IConnector",
            "NetXlsx.INamedRange",
            "NetXlsx.IPicture",
            "NetXlsx.IRange",
            "NetXlsx.IRow",
            "NetXlsx.IShape",
            "NetXlsx.ISheet",
            "NetXlsx.IStreamingCell",
            "NetXlsx.IStreamingRow",
            "NetXlsx.IStreamingSheet",
            "NetXlsx.IStreamingWorkbook",
            "NetXlsx.ITable",
            "NetXlsx.IWorkbook",
            "NetXlsx.IgnoreAttribute",
            "NetXlsx.ImageFormat",
            "NetXlsx.InvalidCellAddressException",
            "NetXlsx.MalformedFileException",
            "NetXlsx.MissingFontException",
            "NetXlsx.NumberFormats",
            "NetXlsx.PictureBorder",
            "NetXlsx.ResourceLimitExceededException",
            "NetXlsx.RichText",
            "NetXlsx.RichTextRun",
            "NetXlsx.RichTextStyle",
            "NetXlsx.ShapeType",
            "NetXlsx.SheetKind",
            "NetXlsx.SheetNameException",
            "NetXlsx.SheetProtection",
            "NetXlsx.SortKey",
            "NetXlsx.StreamingOptions",
            "NetXlsx.StylePoolDiagnostics",
            "NetXlsx.TableStyles",
            "NetXlsx.ThemeColor",
            "NetXlsx.TotalsRowFunction",
            "NetXlsx.UnderlineStyle",
            "NetXlsx.UnsupportedImageFormatException",
            "NetXlsx.VAlign",
            "NetXlsx.Workbook",
            "NetXlsx.WorkbookException",
            "NetXlsx.WorkbookOptions",
            "NetXlsx.WorkbookProtection",
            "NetXlsx.WorksheetAttribute",
            "NetXlsx.WorksheetVisibility",
        };

        publicTypeNames.Should().BeEquivalentTo(expected,
            opt => opt.WithStrictOrdering(),
            "additions must land deliberately via PublicAPI.Unshipped.txt and a test update in the same PR");
    }

    private static bool IsCompilerGenerated(Type t) =>
        t.GetCustomAttributes(typeof(System.Runtime.CompilerServices.CompilerGeneratedAttribute), false).Any();
}
