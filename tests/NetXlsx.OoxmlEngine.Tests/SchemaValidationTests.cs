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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using NetXlsx;
using Xunit;
using ExcelAc = DocumentFormat.OpenXml.Office2013.ExcelAc;
using S = DocumentFormat.OpenXml.Spreadsheet;

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

    // 1×1 transparent PNG, for the drawings fixtures.
    private static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

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

    // ---- Structure: merged regions + named ranges ---------------------------

    [Fact]
    public void Merged_Regions_Are_Schema_Valid()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("Merged");
        s["A1"].SetString("header");
        s.MergeCells("A1:C1");                       // plain merge
        s.MergeCellsStyled("A2:C2", new CellStyle    // styled merge (lesson #4)
        {
            Bold = true,
            Background = Color.FromRgb(0xDD, 0xDD, 0xDD),
            Borders = CellBorders.All(BorderStyle.Thin),
        });
        s.MergeCells("E1:E10");                       // tall merge
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Named_Ranges_WorkbookAnd_SheetScoped_Are_Schema_Valid()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Data");
        wb.AddSheet("Calc");
        wb.AddNamedRange("Global", "Data!$A$1:$A$100");
        wb.AddNamedRange("Scoped", "Calc!$C$3", sheetScope: "Calc");
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- Structure (5b): panes, grouping, visibility, protection ------------

    [Fact]
    public void Sheet_Structure_Surface_Is_Schema_Valid()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Cover");                          // keep >=1 visible sheet
        var s = wb.AddSheet("S");
        for (int i = 1; i <= 6; i++) s.Row(i).Set(1, $"r{i}");
        s.FreezePane(1, 2);                            // frozen panes
        s.GroupRows(2, 5);
        s.GroupRows(3, 4);                             // nested outline
        s.GroupColumns(2, 4);
        s.SetRowGroupCollapsed(3, true);               // collapsed group
        s.ShowGridlines = false;
        s.DefaultColumnWidth = 18;
        s.Hidden = true;
        s.MergeCells("A1:C1");
        s.Protect(password: "pw", options: SheetProtection.LockAll);
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- Drawings (slice 6): pictures ---------------------------------------

    [Fact]
    public void Pictures_Both_Anchor_Kinds_Are_Schema_Valid()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetString("with images");
        s.AddPicture("B2", OnePixelPng, ImageFormat.Png);                 // one-cell anchor
        s.AddPicture("D4", "F8", OnePixelPng, ImageFormat.Png,            // two-cell anchor + EMU offsets
            dx1: 12700, dy1: 6350, dx2: 19050, dy2: 9525);
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- Schema-ordered insertion into OPENED containers (SDK-quirk #8) ------
    //
    // Every fixture above validates a workbook the engine created from scratch,
    // which cannot emit the legal intervening siblings (<autoFilter>,
    // <functionGroups>, …) that sit between a 5a/5b anchor and the element being
    // inserted. These two fixtures close that blind spot: they inject such a
    // sibling at its schema position (simulating an opened real-world file), then
    // mutate through the engine, then validate. A bare InsertAfter(anchor) would
    // emit out-of-order XML here; OoxmlSchemaOrder places the new child correctly.

    [Fact]
    public void Mutating_An_Opened_Sheet_Carrying_AutoFilter_Keeps_Schema_Order()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetString("h");
        s["A2"].SetString("v");

        // <autoFilter> is ubiquitous in real files and sits between <sheetData>
        // and <mergeCells>; the engine does not model it yet, so inject it at its
        // schema position directly.
        var ws = wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        ws.InsertAfter(new S.AutoFilter { Reference = "A1:A2" }, ws.GetFirstChild<S.SheetData>());

        // sheetProtection(before autoFilter) and mergeCells(after autoFilter) must
        // each slot into their schema position around the existing <autoFilter>.
        s.Protect();
        s.MergeCells("A1:A2");
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Mutating_An_Opened_Workbook_Carrying_FunctionGroups_Keeps_Schema_Order()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Data");

        // <functionGroups> sits between <sheets> and <definedNames>; inject it at
        // its schema position to simulate an opened file that carries one.
        var workbook = wb.OpenXmlDocument!.WorkbookPart!.Workbook!;
        workbook.InsertAfter(new S.FunctionGroups(), workbook.GetFirstChild<S.Sheets>());

        wb.AddNamedRange("Items", "Data!$A$1:$A$10");
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Mutating_An_Opened_Workbook_Carrying_AbsPath_Keeps_Schema_Order()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Data");

        // <x15ac:absPath> sits at ordinal 3 (between <workbookPr> and the rest). Excel
        // emits it routinely; its element class lives in an Office2013 extension
        // namespace (DocumentFormat.OpenXml.Office2013.ExcelAc.AbsolutePath), which is
        // why OoxmlSchemaOrder keys by element NAME — the schema sequence is defined
        // over qualified names, not CLR types. This is the workbook analogue of the
        // worksheet's <legacyDrawing> gap, and unlike that one it detonates with a
        // SHIPPING insert: <definedNames> (ordinal 9) would land ahead of an unranked
        // absPath, producing out-of-order XML. Name-keying ranks absPath at 3 so the
        // insert slots in after it.
        var workbook = wb.OpenXmlDocument!.WorkbookPart!.Workbook!;
        workbook.InsertAfter(
            new ExcelAc.AbsolutePath { Url = @"C:\reports\" },
            workbook.GetFirstChild<S.WorkbookProperties>());

        wb.AddNamedRange("Items", "Data!$A$1:$A$10");
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Mutating_An_Opened_Sheet_Carrying_LegacyDrawing_Keeps_Schema_Order()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetString("v");
        s["B1"].SetString("w");

        // <legacyDrawing> (ordinal 29) is the post-<drawing> anchor that real files
        // carry for comments / form controls — ubiquitous, and the reason the post-
        // drawing region matters. Inject one (backed by a VML drawing part so it is
        // schema-valid) to simulate that opened file. Its RANK is machine-checked by
        // SchemaOrderCanonicalTests; behaviorally it only mis-orders once a post-
        // drawing insert (tables / OLE / controls) ships, so this fixture guards that
        // a pre-drawing engine mutation slots in ahead of <legacyDrawing> and leaves
        // the post-drawing region intact.
        var wsp = wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single();
        var ws = wsp.Worksheet!;
        var vml = wsp.AddNewPart<VmlDrawingPart>();
        using (var st = vml.GetStream(FileMode.Create))
        using (var w = new StreamWriter(st))
            w.Write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"></xml>");
        ws.AppendChild(new S.LegacyDrawing { Id = wsp.GetIdOfPart(vml) });

        s.Protect();                 // sheetProtection (ordinal 7)
        s.MergeCells("A1:B1");       // mergeCells (ordinal 14)
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Adding_A_Picture_To_An_Opened_Sheet_Carrying_LegacyDrawing_Keeps_Schema_Order()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetString("v");

        // <legacyDrawing> (ordinal 29) is the FIRST sibling that must follow
        // <drawing> (ordinal 28). Inject one (backed by a VML drawing part so it is
        // schema-valid) to simulate an opened file that carries comments / form
        // controls. AddPicture must slot the new <drawing> in AHEAD of it — a bare
        // AppendChild would emit <drawing> after <legacyDrawing> (out-of-order XML);
        // OoxmlSchemaOrder places it correctly. This is the drawings-slice analogue
        // of the structure slice's open-mutate fixtures (SDK-quirk #8), and the part
        // graph (a new DrawingsPart + ImagePart) makes the opened path matter more
        // here than anywhere prior.
        var wsp = wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single();
        var ws = wsp.Worksheet!;
        var vml = wsp.AddNewPart<VmlDrawingPart>();
        using (var st = vml.GetStream(FileMode.Create))
        using (var w = new StreamWriter(st))
            w.Write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"></xml>");
        ws.AppendChild(new S.LegacyDrawing { Id = wsp.GetIdOfPart(vml) });

        s.AddPicture("B2", "D5", OnePixelPng, ImageFormat.Png);

        // The new <drawing> must precede the pre-existing <legacyDrawing>.
        var children = ws.ChildElements.ToList();
        int drawingIdx = children.FindIndex(c => c is S.Drawing);
        int legacyIdx = children.FindIndex(c => c is S.LegacyDrawing);
        drawingIdx.Should().BeGreaterThanOrEqualTo(0);
        drawingIdx.Should().BeLessThan(legacyIdx);
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- Portability: the gate target's conservative alternative ------------
    //
    // The standing gate target is Microsoft365; design.md records that engine
    // output is also clean under the conservative Office2019 alternative. This
    // guards that documented claim against silent drift — a future slice that
    // emits a Microsoft365-only construct would surface here.

    [Fact]
    public void Created_Output_Also_Validates_Under_Office2019()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetString("text");
        s["A1"].Style(new CellStyle { Bold = true, Background = Color.FromRgb(0xFF, 0xFF, 0x00) });
        s["A2"].SetDate(new DateTime(2026, 5, 31));
        s["A3"].SetRichText(new RichText(
            new RichTextRun("plain", RichTextStyle.Default),
            new RichTextRun(" fancy", new RichTextStyle { Italic = true, Color = Color.FromRgb(0, 0x80, 0) })));
        OpenXmlValidationGate.AssertValid(wb, FileFormatVersions.Office2019);
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
                // Structure slice (I-82): a merged title banner + named ranges,
                // exercised through the open -> resave -> validate path below.
                data["A27"].SetString("Summary");
                data.MergeCellsStyled("A27:B27", new CellStyle { Bold = true, Background = Color.FromRgb(0xEE, 0xEE, 0xEE) });
                wb.AddNamedRange("Items", "Data!$A$2:$A$25");
                wb.AddNamedRange("FirstItem", "Data!$A$2", sheetScope: "Data");

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
