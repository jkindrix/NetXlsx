// NPOI-as-independent-reader oracles — the amendment A1 queued follow-up
// (design.md I-82 cutover note).
//
// At the cutover the cross-engine differential harness was deleted and
// several of its pins became single-engine (SDK-write → SDK-read), which
// is self-referential: a write-side and read-side bug that agree with
// each other pass silently. These tests restore a slim third-party
// verification layer for exactly those surfaces: each writes through the
// default (Open XML SDK) engine and re-opens the saved bytes with RAW
// NPOI — an independent OOXML implementation — asserting the emitted
// markup means what NetXlsx claims it means.
//
// Selection rule (oracle-over-opinion, advisor-confirmed): surfaces whose
// pins went single-engine at the flip (sort, CF, DV — the literal-pinned
// projections), plus surfaces with thin public read-back where an
// independent reader adds real evidence (autofilter + its _FilterDatabase
// name, freeze panes). Scalars/merges/names/annotations/rich text anchor
// the file with the foundational round-trip. NPOI's own read quirks are
// noted inline where they shaped an assertion.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using Xunit;
using NSS = NPOI.SS.UserModel;
using NSU = NPOI.SS.Util;
using NX = NPOI.XSSF.UserModel;

namespace NetXlsx.Tests.Engine;

public class NpoiReaderOracleTests
{
    /// <summary>
    /// Writes a workbook through the default (SDK) engine and re-opens the
    /// saved bytes with raw NPOI. Caller should <c>Close()</c> the result.
    /// </summary>
    private static NX.XSSFWorkbook WriteAndReopen(Action<IWorkbook> author)
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            author(wb);
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        return new NX.XSSFWorkbook(ms);
    }

    [Fact]
    public void Scalars_Formula_And_Date_Read_Back_Via_Npoi()
    {
        var npoi = WriteAndReopen(wb =>
        {
            var s = wb.AddSheet("S");
            s["A1"].SetString("text");
            s["B1"].SetNumber(42.5);
            s["C1"].SetBool(true);
            s["D1"].SetFormula("=SUM(B1:B2)");
            s["E1"].SetDate(new DateTime(2026, 6, 4));
        });

        var row = npoi.GetSheet("S").GetRow(0);
        row.GetCell(0).StringCellValue.Should().Be("text");
        row.GetCell(1).NumericCellValue.Should().Be(42.5);
        row.GetCell(2).BooleanCellValue.Should().BeTrue();
        // NPOI exposes the stored formula body without the leading '='.
        row.GetCell(3).CellFormula.Should().Be("SUM(B1:B2)");
        NSS.DateUtil.IsCellDateFormatted(row.GetCell(4)).Should().BeTrue(
            "the date cell must carry a number format an independent reader classifies as a date");
        row.GetCell(4).DateCellValue.Should().Be(new DateTime(2026, 6, 4));
        npoi.Close();
    }

    [Fact]
    public void Sorted_Range_Reads_Back_In_Order_Via_Npoi()
    {
        // SortTests went single-engine at the flip; the physical row order
        // an independent reader observes is the contract that matters.
        var npoi = WriteAndReopen(wb =>
        {
            var s = wb.AddSheet("S");
            s["A1"].SetString("Charlie");
            s["A2"].SetString("Alice");
            s["A3"].SetString("Bob");
            s.SortRange("A1:A3", SortKey.Asc(1));
        });

        var sheet = npoi.GetSheet("S");
        sheet.GetRow(0).GetCell(0).StringCellValue.Should().Be("Alice");
        sheet.GetRow(1).GetCell(0).StringCellValue.Should().Be("Bob");
        sheet.GetRow(2).GetCell(0).StringCellValue.Should().Be("Charlie");
        npoi.Close();
    }

    [Fact]
    public void Merged_Region_And_Named_Range_Read_Back_Via_Npoi()
    {
        var npoi = WriteAndReopen(wb =>
        {
            var s = wb.AddSheet("Data");
            s["A1"].SetString("anchor");
            s.MergeCells("A1:C3");
            wb.AddNamedRange("Sales", "Data!$A$1:$A$10");
        });

        var sheet = npoi.GetSheet("Data");
        sheet.NumMergedRegions.Should().Be(1);
        sheet.GetMergedRegion(0).FormatAsString().Should().Be("A1:C3");

        var name = npoi.GetName("Sales");
        name.Should().NotBeNull();
        name!.RefersToFormula.Should().Be("Data!$A$1:$A$10");
        npoi.Close();
    }

    [Fact]
    public void Cf_Rule_Reads_Back_Via_Npoi()
    {
        // The CF emission projection was pinned as a LITERAL at the flip;
        // this re-adds an independent reader's interpretation of the same
        // markup (type, operator, formula, range, and the dxf's bold —
        // the exact axis the old projection harness caught a real bug on).
        var npoi = WriteAndReopen(wb =>
        {
            var s = wb.AddSheet("S");
            s.AddConditionalFormatting("A1:A10",
                ConditionalFormat.CellValueGreaterThan("50", new CellStyle { Bold = true }));
        });

        var scf = npoi.GetSheet("S").SheetConditionalFormatting;
        scf.NumConditionalFormattings.Should().Be(1);
        var cf = scf.GetConditionalFormattingAt(0);
        cf.GetFormattingRanges().Single().FormatAsString().Should().Be("A1:A10");
        cf.NumberOfRules.Should().Be(1);

        var rule = cf.GetRule(0);
        rule.ComparisonOperation.Should().Be(NSS.ComparisonOperator.GreaterThan);
        rule.Formula1.Should().Be("50");
        rule.FontFormatting.Should().NotBeNull("the style must land as a dxf with font formatting");
        rule.FontFormatting!.IsBold.Should().BeTrue();
        rule.FontFormatting.IsItalic.Should().BeFalse(
            "a Bold-only CF style must not read back italic — the historical (italic, bold) arg-swap bug class");
        npoi.Close();
    }

    [Fact]
    public void Validation_List_Reads_Back_Via_Npoi()
    {
        // The DV emission projection was pinned as a LITERAL at the flip.
        var npoi = WriteAndReopen(wb =>
        {
            var s = wb.AddSheet("S");
            s.AddValidation("A2:A5", DataValidation.List("Red", "Green", "Blue"));
        });

        var sheet = (NX.XSSFSheet)npoi.GetSheet("S");
        var validations = sheet.GetDataValidations();
        validations.Should().HaveCount(1);

        var dv = validations[0];
        dv.ValidationConstraint.GetValidationType()
            .Should().Be(NSS.ValidationType.LIST);
        dv.ValidationConstraint.Formula1.Should().Be("\"Red,Green,Blue\"");
        dv.Regions.CellRangeAddresses.Single().FormatAsString().Should().Be("A2:A5");
        npoi.Close();
    }

    [Fact]
    public void AutoFilter_Range_And_FilterDatabase_Name_Read_Back_Via_Npoi()
    {
        // The _FilterDatabase semantics were originally oracle-dumped FROM
        // NPOI's emission; reading the SDK engine's output back through
        // NPOI closes that loop from the other side.
        var npoi = WriteAndReopen(wb =>
        {
            var s = wb.AddSheet("S");
            s["A1"].SetString("h1");
            s["B1"].SetString("h2");
            s.SetAutoFilter("A1:B5");
        });

        var sheet = (NX.XSSFSheet)npoi.GetSheet("S");
        sheet.GetCTWorksheet().autoFilter.@ref.Should().Be("A1:B5");

        var filterDb = npoi.GetAllNames()
            .SingleOrDefault(n => n.NameName == "_xlnm._FilterDatabase");
        filterDb.Should().NotBeNull("SetAutoFilter must maintain the hidden built-in name");
        filterDb!.SheetIndex.Should().Be(0);
        npoi.Close();
    }

    [Fact]
    public void Comment_And_Hyperlink_Read_Back_Via_Npoi()
    {
        var npoi = WriteAndReopen(wb =>
        {
            var s = wb.AddSheet("S");
            s["A1"].SetString("flagged");
            s["A1"].Comment("needs review", author: "qa-bot");
            s["B1"].SetString("link");
            s["B1"].Hyperlink("https://example.com/docs");
        });

        var sheet = npoi.GetSheet("S");
        var comment = sheet.GetCellComment(new NSU.CellAddress(0, 0));
        comment.Should().NotBeNull();
        comment!.String.String.Should().Be("needs review");
        comment.Author.Should().Be("qa-bot");

        var link = sheet.GetRow(0).GetCell(1).Hyperlink;
        link.Should().NotBeNull();
        link!.Address.Should().Be("https://example.com/docs");
        npoi.Close();
    }

    [Fact]
    public void Freeze_Pane_Reads_Back_Via_Npoi()
    {
        var npoi = WriteAndReopen(wb =>
        {
            var s = wb.AddSheet("S");
            s["A1"].SetString("header");
            s.FreezeRows(1);
        });

        var pane = npoi.GetSheet("S").PaneInformation;
        pane.Should().NotBeNull();
        pane!.IsFreezePane().Should().BeTrue();
        pane.HorizontalSplitPosition.Should().Be(1);
        pane.VerticalSplitPosition.Should().Be(0);
        npoi.Close();
    }

    [Fact]
    public void RichText_Inheriting_Run_Reads_Back_Fontless_Via_Npoi()
    {
        // The marquee rich-text win (lesson #10 / the N1 fidelity fix): a
        // run with no <rPr> inherits the cell font. NPOI reads such a run
        // with a null font — independent confirmation the inheritance
        // semantic survived the engine's emission.
        var npoi = WriteAndReopen(wb =>
        {
            var s = wb.AddSheet("S");
            s["A1"].SetRichText(new RichText(
                new RichTextRun("VERY IMPORTANT: "),
                new RichTextRun("read the notes", new RichTextStyle { Bold = true })));
        });

        var rich = (NX.XSSFRichTextString)npoi.GetSheet("S").GetRow(0).GetCell(0).RichStringCellValue;
        rich.String.Should().Be("VERY IMPORTANT: read the notes");
        rich.NumFormattingRuns.Should().Be(2);
        rich.GetFontOfFormattingRun(0).Should().BeNull(
            "the unstyled prefix run must carry no <rPr> so it inherits the cell font");
        rich.GetFontOfFormattingRun(1).IsBold.Should().BeTrue();
        npoi.Close();
    }
}
