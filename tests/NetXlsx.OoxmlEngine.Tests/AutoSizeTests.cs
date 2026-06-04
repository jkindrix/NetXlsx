// I-84 — IColumn.AutoSize conformance on the Open XML SDK engine.
//
// AutoSize measures with EMBEDDED font-metric tables (OoxmlFontMetrics /
// FontMetricsTables.g.cs), so unlike the NPOI engine's environment-dependent
// AutoSize (I-3), every width here is DETERMINISTIC: same input, same result,
// on any machine, headless or not. That is why these tests pin exact expected
// widths — they are properties of the embedded tables + the NPOI-mirrored
// SheetUtil formula, not of the host's font stack.
//
// The formula (oracle-dumped from NPOI 2.7.3, 2026-06-03):
//   cellWidth = (round(linePx, ToEven) + 5) / defaultCharWidth * 1.05 + indent
//   defaultCharWidth = ceil(ink('0') of font 0 @96dpi) — 7 for Calibri 11.
//
// Documented divergences from the NPOI engine (design.md I-84):
//   - fonts outside the embedded set throw MissingFontException (NPOI silently
//     falls back to Arial / the first system family);
//   - a fresh formula with no cached result is skipped (NPOI measures the
//     literal "0" its empty cached <v/> reads back as);
//   - non-date custom number formats measure shortest round-trip text
//     (NPOI renders through DataFormatter, e.g. thousands separators).

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using S = DocumentFormat.OpenXml.Spreadsheet;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class AutoSizeTests
{
    private const string LongText = "Hello, World! A longer string of text";

    private static double AutoSizedWidth(Action<ISheet> populate, string column = "A")
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        populate(sheet);
        return sheet.Column(column).AutoSize().WidthUnits;
    }

    // ---- pinned deterministic widths ---------------------------------------

    [Fact]
    public void String_Default_Font_Pins_Width()
        => AutoSizedWidth(s => s["A1"].SetString(LongText))
            .Should().BeApproximately(32.55, 1e-9);

    [Fact]
    public void Bold_Pins_Width()
        => AutoSizedWidth(s =>
        {
            s["A1"].SetString(LongText);
            s["A1"].Style(new CellStyle { Bold = true });
        }).Should().BeApproximately(33.30, 1e-9);

    [Fact]
    public void Arial_10_Pins_Width()
        => AutoSizedWidth(s =>
        {
            s["A1"].SetString(LongText);
            s["A1"].Style(new CellStyle { FontName = "Arial", FontSize = 10 });
        }).Should().BeApproximately(31.65, 1e-9);

    [Fact]
    public void Number_General_Pins_Width()
        => AutoSizedWidth(s => s["A1"].SetNumber(1234.5678))
            .Should().BeApproximately(10.2, 1e-9);

    [Fact]
    public void Date_Renders_Through_Format_And_Pins_Width()
        => AutoSizedWidth(s =>
        {
            s["A1"].SetDate(new DateTime(2026, 6, 3));
            s["A1"].NumberFormat("d-mmm-yy");
        }).Should().BeApproximately(8.4, 1e-9);

    [Fact]
    public void Bool_Measures_TRUE_And_Pins_Width()
        => AutoSizedWidth(s => s["A1"].SetBool(true))
            .Should().BeApproximately(5.55, 1e-9);

    [Fact]
    public void Multiline_Measures_Longest_Line()
    {
        double multiline = AutoSizedWidth(s => s["A1"].SetString("abc\nlonger line here"));
        double longestOnly = AutoSizedWidth(s => s["A1"].SetString("longer line here"));
        multiline.Should().BeApproximately(longestOnly, 1e-12);
        multiline.Should().BeApproximately(14.85, 1e-9);
    }

    [Fact]
    public void Widest_Row_Wins()
        => AutoSizedWidth(s =>
        {
            s["A1"].SetString("abc");
            s["A2"].SetString("the longest row in this column wins");
        }).Should().BeApproximately(32.4, 1e-9);

    [Fact]
    public void Date_Cell_Width_Equals_Equivalent_String_Width()
    {
        // The date cell measures its rendered text — identical text in the
        // same font must produce the identical width.
        double date = AutoSizedWidth(s =>
        {
            s["A1"].SetDate(new DateTime(2026, 6, 3));
            s["A1"].NumberFormat("d-mmm-yy");
        });
        double str = AutoSizedWidth(s => s["A1"].SetString("3-Jun-26"));
        date.Should().BeApproximately(str, 1e-12);
    }

    // ---- skip semantics (NPOI AutoSizeColumn parity) -----------------------

    [Fact]
    public void Empty_Column_Is_A_NoOp()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        var col = sheet.Column("A");

        col.AutoSize().Should().BeSameAs(col);   // fluent chaining
        col.WidthUnits.Should().BeApproximately(8.43, 1e-9);   // untouched default

        // No <col> element materializes — NPOI's AutoSizeColumn does nothing
        // when every cell is empty, and so does this engine.
        wb.Underlying.WorkbookPart!.WorksheetParts.Single().Worksheet!
            .GetFirstChild<S.Columns>().Should().BeNull();
    }

    [Fact]
    public void Merged_Cells_Are_Skipped()
    {
        // NPOI parity: AutoSizeColumn(col) runs with useMergedCells=false —
        // a column whose only content sits in a merged region is a no-op.
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet.MergeCells("A1:B2");
        sheet["A1"].SetString("Merged content string");

        sheet.Column("A").AutoSize().WidthUnits.Should().BeApproximately(8.43, 1e-9);
    }

    [Fact]
    public void Merged_Cell_Skipped_But_Standalone_Cell_Still_Measured()
    {
        double mixed = AutoSizedWidth(s =>
        {
            s.MergeCells("A1:B2");
            s["A1"].SetString("Merged content string that is quite long indeed");
            s["A3"].SetString("standalone");
        });
        double standaloneOnly = AutoSizedWidth(s => s["A1"].SetString("standalone"));
        mixed.Should().BeApproximately(standaloneOnly, 1e-12);
    }

    [Fact]
    public void Fresh_Formula_Without_Cached_Result_Is_Skipped()
    {
        // Documented I-84 divergence: NPOI measures the "0" its empty cached
        // <v/> reads back as (an artifact); this engine skips the cell.
        AutoSizedWidth(s => s["A1"].SetFormula("=1+2"))
            .Should().BeApproximately(8.43, 1e-9);
    }

    [Fact]
    public void Formula_With_Cached_Result_Measures_The_Cached_Text()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetFormula("=1+2");

        // Files opened from disk can carry cached formula results; emulate one
        // through the escape hatch.
        var cell = wb.Underlying.WorkbookPart!.WorksheetParts.Single()
            .Worksheet!.Descendants<S.Cell>().Single(c => c.CellReference!.Value == "A1");
        cell.AppendChild(new S.CellValue("1234.5678"));

        double formula = sheet.Column("A").AutoSize().WidthUnits;
        double number = AutoSizedWidth(s => s["A1"].SetNumber(1234.5678));
        formula.Should().BeApproximately(number, 1e-12);
    }

    // ---- fail-loud font contract (I-84, superseding I-3) -------------------

    [Fact]
    public void Unknown_Font_Throws_MissingFontException()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("text");
        sheet["A1"].Style(new CellStyle { FontName = "NoSuchFont123" });

        Action act = () => sheet.Column("A").AutoSize();
        act.Should().Throw<MissingFontException>()
            .Which.Message.Should().Contain("NoSuchFont123")
            .And.Contain("IColumn.Width", "the message must point at the explicit-width alternative");
    }

    [Fact]
    public void Unknown_Font_On_Empty_Column_Does_Not_Throw()
    {
        // Fonts resolve only for cells that are actually measured — an empty
        // column never measures, so it never throws.
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["B1"].SetString("other column");
        sheet["B1"].Style(new CellStyle { FontName = "NoSuchFont123" });

        ((Action)(() => sheet.Column("A").AutoSize())).Should().NotThrow();
    }

    [Theory]
    [InlineData("Calibri", "Carlito")]
    [InlineData("Calibri", "Aptos")]            // best-effort stand-in (I-84)
    [InlineData("Calibri", "Aptos Narrow")]
    [InlineData("Arial", "Liberation Sans")]
    [InlineData("Arial", "Arimo")]
    [InlineData("Arial", "Helvetica")]
    [InlineData("Times New Roman", "Liberation Serif")]
    [InlineData("Times New Roman", "Tinos")]
    [InlineData("Courier New", "Liberation Mono")]
    [InlineData("Courier New", "Cousine")]
    public void Metric_Twin_Names_Produce_Identical_Widths(string canonical, string twin)
    {
        double Width(string fontName) => AutoSizedWidth(s =>
        {
            s["A1"].SetString(LongText);
            s["A1"].Style(new CellStyle { FontName = fontName, FontSize = 12 });
        });
        Width(twin).Should().BeApproximately(Width(canonical), 1e-12);
    }

    [Fact]
    public void Unmapped_Char_Falls_Back_To_Digit_Advance()
    {
        // A char with no table entry (here CJK) measures as the digit advance,
        // so a string of N unmapped chars is exactly as wide as N digits.
        double cjk = AutoSizedWidth(s => s["A1"].SetString("中中中"));
        double digits = AutoSizedWidth(s => s["A1"].SetString("000"));
        cjk.Should().BeApproximately(digits, 1e-12);
    }

    [Fact]
    public void Typographic_Extras_Have_Real_Metrics()
    {
        // Em-dash and friends live in the extras table, not the digit fallback:
        // an em-dash (nearly 1em wide) must measure wider than a digit.
        double emDashes = AutoSizedWidth(s => s["A1"].SetString("———"));
        double digits = AutoSizedWidth(s => s["A1"].SetString("000"));
        emDashes.Should().BeGreaterThan(digits);
    }

    // ---- output shape -------------------------------------------------------

    [Fact]
    public void Width_Caps_At_255_Units()
        => AutoSizedWidth(s => s["A1"].SetString(new string('W', 400)))
            .Should().Be(255);

    [Fact]
    public void Col_Element_Carries_BestFit_And_CustomWidth()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("some content");
        sheet.Column("A").AutoSize();

        var col = wb.Underlying.WorkbookPart!.WorksheetParts.Single()
            .Worksheet!.GetFirstChild<S.Columns>()!.Elements<S.Column>().Single();
        col.BestFit!.Value.Should().BeTrue();
        col.CustomWidth!.Value.Should().BeTrue();

        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void AutoSized_Width_RoundTrips_Through_Save_And_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-autosize-{Guid.NewGuid():N}.xlsx");
        try
        {
            double saved;
            using (var wb = Workbook.CreateOoxml())
            {
                var sheet = wb.AddSheet("S");
                sheet["A1"].SetString(LongText);
                saved = sheet.Column("A").AutoSize().WidthUnits;
                wb.Save(path);
            }
            using var reopened = Workbook.OpenOoxml(path);
            reopened["S"].Column("A").WidthUnits.Should().BeApproximately(saved, 1e-12);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- opened-file style axes (indent / rotation) -------------------------

    [Fact]
    public void Alignment_Indent_Adds_Width_Units()
    {
        // CellStyle cannot author indent, but opened files carry it; emulate
        // through the escape hatch. NPOI adds the raw indent level AFTER the
        // correction multiply — so the delta is exactly the indent value.
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString(LongText);
        sheet["A1"].Style(new CellStyle { Bold = true });   // forces a non-zero xf

        var ss = wb.Underlying.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        var xf = ss.GetFirstChild<S.CellFormats>()!.Elements<S.CellFormat>().Last();
        xf.Alignment = new S.Alignment { Indent = 2 };
        xf.ApplyAlignment = true;

        double indented = sheet.Column("A").AutoSize().WidthUnits;
        double plain = AutoSizedWidth(s =>
        {
            s["A1"].SetString(LongText);
            s["A1"].Style(new CellStyle { Bold = true });
        });
        indented.Should().BeApproximately(plain + 2, 1e-9);
    }

    [Fact]
    public void Vertical_Rotation_Narrows_A_Long_String_Column()
    {
        // textRotation=90 turns the long axis vertical: cos(90°)=0, so only
        // the (approximated) line height contributes — far narrower than the
        // unrotated measurement of the same string.
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString(LongText);
        sheet["A1"].Style(new CellStyle { Bold = true });

        var ss = wb.Underlying.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        var xf = ss.GetFirstChild<S.CellFormats>()!.Elements<S.CellFormat>().Last();
        xf.Alignment = new S.Alignment { TextRotation = 90 };
        xf.ApplyAlignment = true;

        double rotated = sheet.Column("A").AutoSize().WidthUnits;
        double unrotated = AutoSizedWidth(s =>
        {
            s["A1"].SetString(LongText);
            s["A1"].Style(new CellStyle { Bold = true });
        });
        rotated.Should().BeLessThan(unrotated / 2);
        rotated.Should().BeGreaterThan(0);
    }

    // ---- lifecycle ----------------------------------------------------------

    [Fact]
    public void AutoSize_On_Disposed_Workbook_Throws()
    {
        var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("x");
        var col = sheet.Column("A");
        wb.Dispose();

        ((Action)(() => col.AutoSize())).Should().Throw<ObjectDisposedException>();
    }
}
