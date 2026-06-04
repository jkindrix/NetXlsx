// I-82 engine swap — CF/validation/tables/autofilter/sort slice: SortRange
// conformance.
//
// Mirrors the NPOI engine's SortTests contract on the Open XML SDK engine
// (decision I-72): single/multi key ascending + descending, stability for
// tied rows (lesson #12 — Excel's sort is stable), styles moving with their
// values, blanks sorting last, the 1-row no-op, argument validation, and a
// Save/Open round-trip. SDK-engine extras: rich-text cells move with their
// runs intact (the engine moves the <c> element, not a value snapshot), and
// formula cells from an OPENED file keep their literal text in their new row
// (lesson #12 — references are NOT relocated).

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.OoxmlEngine.Tests;

public class SortTests
{
    [Fact]
    public void SortRange_Ascending_String_Column()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetString("Charlie");
        s["A2"].SetString("Alice");
        s["A3"].SetString("Bob");

        s.SortRange("A1:A3", SortKey.Asc(1));

        s["A1"].GetString().Should().Be("Alice");
        s["A2"].GetString().Should().Be("Bob");
        s["A3"].GetString().Should().Be("Charlie");
    }

    [Fact]
    public void SortRange_Descending_Numeric_Column()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetNumber(10);
        s["A2"].SetNumber(30);
        s["A3"].SetNumber(20);

        s.SortRange("A1:A3", SortKey.Desc(1));

        s["A1"].GetNumber().Should().Be(30);
        s["A2"].GetNumber().Should().Be(20);
        s["A3"].GetNumber().Should().Be(10);
    }

    [Fact]
    public void SortRange_Multi_Column_Key()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetString("B"); s["B1"].SetNumber(2);
        s["A2"].SetString("A"); s["B2"].SetNumber(3);
        s["A3"].SetString("A"); s["B3"].SetNumber(1);

        s.SortRange("A1:B3", SortKey.Asc(1), SortKey.Asc(2));

        s["A1"].GetString().Should().Be("A");
        s["B1"].GetNumber().Should().Be(1);
        s["A2"].GetString().Should().Be("A");
        s["B2"].GetNumber().Should().Be(3);
        s["A3"].GetString().Should().Be("B");
        s["B3"].GetNumber().Should().Be(2);
    }

    [Fact]
    public void SortRange_Is_Stable_For_Tied_Rows()
    {
        // Rows tying on the sort key must keep their original relative
        // order (Excel's sort is stable). Column B is a witness to the
        // original order; it is not a sort key.
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetString("B"); s["B1"].SetNumber(1);
        s["A2"].SetString("A"); s["B2"].SetNumber(2);
        s["A3"].SetString("A"); s["B3"].SetNumber(3);
        s["A4"].SetString("A"); s["B4"].SetNumber(4);

        s.SortRange("A1:B4", SortKey.Asc(1));

        // The three "A" rows keep original order (witnesses 2,3,4), then "B".
        s["B1"].GetNumber().Should().Be(2);
        s["B2"].GetNumber().Should().Be(3);
        s["B3"].GetNumber().Should().Be(4);
        s["B4"].GetNumber().Should().Be(1);
    }

    [Fact]
    public void SortRange_Preserves_Styles()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        var bold = new CellStyle { Bold = true };
        s["A1"].SetString("B"); s["A1"].Style(bold);
        s["A2"].SetString("A");

        s.SortRange("A1:A2", SortKey.Asc(1));

        s["A1"].GetString().Should().Be("A");
        s["A2"].GetString().Should().Be("B");
        // The bold style moved with "B"
        s["A2"].GetStyle().Bold.Should().BeTrue();
    }

    [Fact]
    public void SortRange_Blanks_Sort_Last_Ascending()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A2"].SetNumber(1);
        s["A3"].SetNumber(3);
        // A1 is blank

        s.SortRange("A1:A3", SortKey.Asc(1));

        s["A1"].GetNumber().Should().Be(1);
        s["A2"].GetNumber().Should().Be(3);
        s["A3"].Kind.Should().Be(CellKind.Empty);
    }

    [Fact]
    public void SortRange_Single_Row_Is_Noop()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetString("only");

        s.SortRange("A1:A1", SortKey.Asc(1));

        s["A1"].GetString().Should().Be("only");
    }

    [Fact]
    public void SortRange_Rejects_Null_Range()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        Action act = () => s.SortRange(null!, SortKey.Asc(1));
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void SortRange_Rejects_Empty_Keys()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        Action act = () => s.SortRange("A1:A3");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void SortRange_Survives_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.CreateOoxml())
        {
            var s = wb.AddSheet("S");
            s["A1"].SetNumber(30);
            s["A2"].SetNumber(10);
            s["A3"].SetNumber(20);
            s.SortRange("A1:A3", SortKey.Asc(1));
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.OpenOoxml(ms);
        opened["S"]["A1"].GetNumber().Should().Be(10);
        opened["S"]["A2"].GetNumber().Should().Be(20);
        opened["S"]["A3"].GetNumber().Should().Be(30);
    }

    [Fact]
    public void SortRange_Moves_RichText_With_Its_Runs()
    {
        // The SDK engine moves the <c> element itself, so an inline rich
        // string's runs (including run-level formatting) travel with the row.
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetRichText(new RichText(
            new RichTextRun("zebra ", RichTextStyle.Default),
            new RichTextRun("tail", new RichTextStyle { Bold = true })));
        s["A2"].SetString("apple");

        s.SortRange("A1:A2", SortKey.Asc(1));

        s["A1"].GetString().Should().Be("apple");
        var rich = s["A2"].GetRichText();
        rich.Should().NotBeNull();
        rich!.Runs.Should().HaveCount(2);
        rich.Runs[0].Text.Should().Be("zebra ");
        rich.Runs[1].Text.Should().Be("tail");
        rich.Runs[1].Style.Bold.Should().BeTrue();
    }

    [Fact]
    public void SortRange_Moves_Formula_Text_Verbatim()
    {
        // Lesson #12 / the documented formula caveat: a sorted formula cell
        // keeps its literal text — references are NOT relocated. Round-trips
        // through Save/Open so the sort runs on an OPENED file, not just the
        // authoring DOM.
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-sort-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var author = Workbook.Create())
            {
                var sheet = author.AddSheet("S");
                sheet["A1"].SetString("zulu");
                sheet["B1"].SetFormula("A1&\"!\"");
                sheet["A2"].SetString("alpha");
                sheet["B2"].SetNumber(7);
                author.Save(path);
            }

            using var wb = Workbook.Open(path);
            var s = wb["S"];
            s.SortRange("A1:B2", SortKey.Asc(1));

            s["A1"].GetString().Should().Be("alpha");
            s["A2"].GetString().Should().Be("zulu");
            // The formula moved to row 2 with its literal text intact.
            s["B2"].Kind.Should().Be(CellKind.Formula);
            var ws = wb.Underlying.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var b2 = ws.Descendants<S.Cell>().Single(c => c.CellReference?.Value == "B2");
            b2.CellFormula!.Text.Should().Be("A1&\"!\"");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SortRange_Moves_A_Styled_Blank_Cell()
    {
        // A materialized-but-blank cell (<c s="N"/>) is Empty for comparison
        // (blanks sort last) but its element — and so its style — still moves.
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].Style(new CellStyle { Bold = true }); // styled, no value
        s["A2"].SetString("value");

        s.SortRange("A1:A2", SortKey.Asc(1));

        s["A1"].GetString().Should().Be("value");
        s["A2"].Kind.Should().Be(CellKind.Empty);
        s["A2"].GetStyle().Bold.Should().BeTrue();
    }
}
