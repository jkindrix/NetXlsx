// I-82 engine swap — schema-validation conformance gate (slice 4b).
//
// The founding premise of the engine swap is that the Open XML SDK is
// schema-complete: correctness "for free." Nothing previously validated the
// engine's OUTPUT against that schema — Excel rendering an element is not proof
// of schema-validity. This gate runs DocumentFormat.OpenXml's OpenXmlValidator
// over engine output and asserts zero errors, so every landed feature (and every
// future slice that reuses these fixtures) is held to the schema, not to "Excel
// opened it without complaining."
//
// Validation target — Microsoft365 (I-82 sub-decision, recorded in design.md
// §6.2.15). Microsoft365 is the most current FileFormatVersions; the engine
// targets modern Excel and round-trips Microsoft365-era parts (e.g. x14/x15
// extensions) unmodeled. All created-workbook fixtures below also validate clean
// under the conservative Office2019 alternative (checked while authoring this
// gate); Microsoft365 is the standing gate target.
//
// What this gate established about the rich-text <rPr> child order — the prime
// suspect carried in the handoff notes: the SDK validator does NOT constrain
// CT_RPrElt child order (current order, strict-ECMA order, and a deliberately
// scrambled order all validate clean), whereas it DOES constrain CT_Font order
// in styles.xml (a scrambled <font> raises Sch_UnexpectedElementContentExpecting
// Complex). So the engine's <rPr> emit order is schema-valid as-is — there was
// nothing to reorder. The font path's order was already correct. See SDK-quirk #5.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

/// <summary>
/// Reusable schema-validation gate for the Open XML SDK engine. Runs
/// <see cref="OpenXmlValidator"/> over a workbook's live
/// <see cref="IWorkbook.OpenXmlDocument"/> and fails with a fully-itemized
/// dump (part URI, element XPath, error id, description) when the schema is
/// violated. Future slices should call <see cref="AssertValid"/> on their own
/// fixtures so the gate widens with the engine.
/// </summary>
public static class OpenXmlValidationGate
{
    /// <summary>The standing validation target — see the file header / design.md §6.2.15 (I-82).</summary>
    public const FileFormatVersions Target = FileFormatVersions.Microsoft365;

    public static void AssertValid(IWorkbook workbook, FileFormatVersions? version = null)
    {
        var doc = workbook.OpenXmlDocument
            ?? throw new InvalidOperationException(
                "AssertValid requires the SDK engine (CreateOoxml/OpenOoxml); OpenXmlDocument was null.");

        var resolved = version ?? Target;
        var validator = new OpenXmlValidator(resolved);
        var errors = validator.Validate(doc).ToList();
        if (errors.Count == 0) return;

        // Itemize each error (part URI / element XPath / id / description) so a
        // failure is diagnosable from the assertion message alone. Built with
        // string.Join — interpolated AppendLine trips CA1305 (locale-sensitive).
        var detail = string.Join(Environment.NewLine, errors.Select(e =>
            $"  [{e.ErrorType}] {e.Id}{Environment.NewLine}" +
            $"    part: {e.Part?.Uri?.ToString() ?? "(none)"}{Environment.NewLine}" +
            $"    path: {e.Path?.XPath ?? "(none)"}{Environment.NewLine}" +
            $"    node: {e.Node?.LocalName ?? "(none)"}{Environment.NewLine}" +
            $"    desc: {e.Description}"));
        errors.Should().BeEmpty(
            $"the workbook must be schema-valid ({resolved}); validator reported "
            + $"{errors.Count} error(s):{Environment.NewLine}{detail}");
    }
}

public class SchemaValidationTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-validate-{Guid.NewGuid():N}.xlsx");

    // ---- Values: string / number / bool -------------------------------------

    [Fact]
    public void Created_Workbook_With_Scalar_Values_Is_Schema_Valid()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("Vals");
        s["A1"].SetString("hello");
        s["B1"].SetNumber(42.5);
        s["C1"].SetBool(true);
        s["A2"].SetNumber(-1234567);
        s["B2"].SetString("");          // empty string
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- Styles: fonts / fills / borders / numFmts / dates ------------------

    [Fact]
    public void Styled_Cells_Are_Schema_Valid()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("Styled");

        s["A1"].SetString("bold");
        s["A1"].Style(new CellStyle
        {
            Bold = true,
            Italic = true,
            Underline = UnderlineStyle.Single,
            FontName = "Arial",
            FontSize = 14,
            FontColor = Color.FromRgb(0x10, 0x20, 0x30),
        });

        s["A2"].SetNumber(1);
        s["A2"].Style(new CellStyle { Background = Color.FromRgb(0xFF, 0xFF, 0x00) });

        s["A3"].SetNumber(2);
        s["A3"].Style(new CellStyle
        {
            Borders = CellBorders.All(BorderStyle.Thin),
            HorizontalAlignment = HAlign.Center,
            VerticalAlignment = VAlign.Center,
            WrapText = true,
        });

        s["A4"].SetNumber(1234.56);
        s["A4"].NumberFormat("#,##0.00");

        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Date_And_Time_Cells_Are_Schema_Valid()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("Dates");
        s["A1"].SetDate(new DateTime(2026, 5, 31));
        s["A2"].SetTime(new TimeOnly(13, 44, 59));
        s["A3"].SetDuration(TimeSpan.FromHours(36.5));
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Date_Cells_Under_The_1904_Epoch_Are_Schema_Valid()
    {
        using var wb = Workbook.CreateOoxml(new WorkbookOptions { DateSystem = DateSystem.Excel1904 });
        wb.AddSheet("D")["A1"].SetDate(new DateTime(2026, 5, 31));
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- Rich text, incl. the empty-style inheriting run --------------------

    [Fact]
    public void Rich_Text_Including_An_Inheriting_Run_Is_Schema_Valid()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Rich")["A1"].SetRichText(new RichText(
            // Empty-style run: emitted with NO <rPr> (inherits cell font, lesson #10).
            new RichTextRun("VERY IMPORTANT", RichTextStyle.Default),
            // Fully-formatted run: exercises every <rPr> axis the engine emits.
            new RichTextRun(" please read", new RichTextStyle
            {
                Bold = true,
                Italic = true,
                Underline = UnderlineStyle.Single,
                FontName = "Arial",
                FontSize = 14,
                Color = Color.FromRgb(0xFF, 0, 0),
            })));
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- OpenOoxml -> Save round-trip ---------------------------------------
    //
    // CI-safe stand-in for a "real stress file" round-trip: the project commits no
    // binary .xlsx fixtures (decision I18 option b — fixtures are built on demand,
    // never blobbed), and the five real stress pairs live only in the operator's
    // Downloads, which CI cannot reach. So this builds a workbook spanning every
    // landed feature across several sheets, saves it, reopens via OpenOoxml, saves
    // AGAIN, and validates the reopened package — exercising the full
    // open-existing-package-and-resave path under the schema gate. (The five real
    // stress files were validated manually while authoring this gate: four are
    // clean; ANIMAL_STRAW_HOLDERS_PSS carries one pre-existing source-authored
    // x14:workbookPr/@defaultImageDpi='32767' that the engine OPC-preserves
    // verbatim per lesson #13 — not engine-generated, and correct to preserve.)
    //
    // Validation cost: sub-300 ms even on the 3.9 MB ANIMAL file; this synthetic
    // round-trip validates in single-digit ms.

    [Fact]
    public void OpenOoxml_Save_RoundTrip_Of_A_Rich_Workbook_Is_Schema_Valid()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var data = wb.AddSheet("Data");
                data["A1"].SetString("name");
                data["B1"].SetString("qty");
                data["A1"].Style(new CellStyle { Bold = true, Background = Color.FromRgb(0xDD, 0xDD, 0xDD) });
                data["B1"].Style(new CellStyle { Bold = true, Background = Color.FromRgb(0xDD, 0xDD, 0xDD) });
                for (int r = 2; r <= 25; r++)
                {
                    data[$"A{r}"].SetString($"item-{r}");
                    data[$"B{r}"].SetNumber(r * 1.5);
                    data[$"B{r}"].NumberFormat("#,##0.00");
                }

                var meta = wb.AddSheet("Meta");
                meta["A1"].SetRichText(new RichText(
                    new RichTextRun("Report", RichTextStyle.Default),
                    new RichTextRun(" (draft)", new RichTextStyle { Italic = true, Color = Color.FromRgb(0x80, 0x80, 0x80) })));
                meta["A2"].SetDate(new DateTime(2026, 5, 31));
                meta["A3"].SetBool(true);
                meta["A4"].Style(new CellStyle { Borders = CellBorders.All(BorderStyle.Medium) });
                meta.Column("A").Width(24);

                wb.Save(path);
            }

            using (var wb = Workbook.OpenOoxml(path))
            {
                // Resave to exercise the open-existing -> mutate-nothing -> save path,
                // then validate the live package.
                var resaved = TempXlsxPath();
                try
                {
                    wb.Save(resaved);
                    OpenXmlValidationGate.AssertValid(wb);
                }
                finally { if (File.Exists(resaved)) File.Delete(resaved); }
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
