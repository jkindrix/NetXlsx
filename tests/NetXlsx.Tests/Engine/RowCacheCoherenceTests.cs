// I-87 — row-index cache coherence on the SDK engine.
//
// The v2.0.1 perf fix caches row-index -> <row> element lookups on OoxmlSheet
// (the pre-fix linear scans made bulk DOM writes O(n²)). The cache must be
// invisible: every mutation route to the worksheet DOM — engine APIs, the
// escape hatches, and even stored hatch references mutated out-of-contract —
// must leave facade reads correct.
//
// Coherence model under test (the I-87 design):
//   1. Every Underlying getter (workbook / sheet / row / cell) invalidates the
//      row caches, so acquire-then-mutate-then-facade is always coherent.
//   2. Backstop liveness checks: a cached row that was removed or renumbered
//      through a STORED reference (no re-acquire) is detected per hit and the
//      cache rebuilds rather than serving a detached node.
//   3. Tail verification: a row appended through a stored reference is
//      detected by AppendRow/MaxRowIndex's O(1) live-tail check.
//
// These tests deliberately reach through .Underlying to mutate the DOM the
// way an escape-hatch consumer would.

using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.Tests.Engine;

public class RowCacheCoherenceTests
{
    private static S.SheetData SheetDataOf(ISheet sheet)
        => sheet.Underlying.GetFirstChild<S.SheetData>()!;

    // ---- Contract pattern: acquire hatch, mutate, continue via facade ------

    [Fact]
    public void Hatch_Removed_Row_Reads_As_Empty()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        for (int r = 1; r <= 10; r++) sheet.Row(r).Set(1, $"row-{r}");

        var data = SheetDataOf(sheet); // hatch access — invalidates
        data.Elements<S.Row>().Single(r => r.RowIndex!.Value == 5u).Remove();

        sheet["A5"].Kind.Should().Be(CellKind.Empty);
        sheet["A5"].GetString().Should().BeEmpty();
        sheet["A4"].GetString().Should().Be("row-4");
        sheet["A6"].GetString().Should().Be("row-6");
        sheet.LastRowNumber.Should().Be(10);
    }

    [Fact]
    public void Hatch_Renumbered_Row_Is_Read_At_Its_New_Index()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        // Height-only rows (no cells) so renumbering @r needs no cell-ref fixup.
        sheet.Row(5).HeightInPoints = 33f;

        var data = SheetDataOf(sheet); // hatch access — invalidates
        data.Elements<S.Row>().Single(r => r.RowIndex!.Value == 5u).RowIndex = 7u;

        sheet.Row(7).HeightInPoints.Should().Be(33f);
        sheet.Row(5).HeightInPoints.Should().Be(15f); // back to default
    }

    [Fact]
    public void Hatch_Added_Row_Is_Visible_And_AppendRow_Lands_After_It()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        for (int r = 1; r <= 3; r++) sheet.Row(r).Set(1, r);

        var data = SheetDataOf(sheet); // hatch access — invalidates
        data.AppendChild(new S.Row { RowIndex = 11u, Height = 42d, CustomHeight = true });

        sheet.Row(11).HeightInPoints.Should().Be(42f);
        var appended = sheet.AppendRow();
        appended.Set(1, "tail");
        appended.Index.Should().Be(12);
        sheet["A12"].GetString().Should().Be("tail");
    }

    [Fact]
    public void Workbook_Hatch_Invalidates_Every_Sheet()
    {
        using var wb = Workbook.Create();
        var one = wb.AddSheet("One");
        var two = wb.AddSheet("Two");
        one.Row(1).Set(1, "a");
        two.Row(1).Set(1, "b");

        // Mutate BOTH sheets through the workbook-level hatch.
        var doc = wb.Underlying;
        foreach (var wsPart in doc.WorkbookPart!.WorksheetParts)
            wsPart.Worksheet!.GetFirstChild<S.SheetData>()!
                .Elements<S.Row>().Single().Remove();

        one["A1"].Kind.Should().Be(CellKind.Empty);
        two["A1"].Kind.Should().Be(CellKind.Empty);
    }

    [Fact]
    public void Row_And_Cell_Hatches_Invalidate_Too()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        sheet.Row(1).Set(1, "x").Set(2, "y");

        // IRow.Underlying: strip the row's cells through the row hatch.
        sheet.Row(1).Underlying.RemoveAllChildren<S.Cell>();
        sheet["A1"].Kind.Should().Be(CellKind.Empty);
        sheet["B1"].Kind.Should().Be(CellKind.Empty);

        // ICell.Underlying: remove the parent row through the cell hatch.
        sheet.Row(2).Set(1, "z");
        var cell = sheet["A2"].Underlying;
        cell.Parent!.Remove();
        sheet["A2"].Kind.Should().Be(CellKind.Empty);
    }

    // ---- Backstop: stored reference mutated AFTER intervening facade calls --

    [Fact]
    public void Stored_Reference_Removal_Is_Caught_By_The_Liveness_Check()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        for (int r = 1; r <= 5; r++) sheet.Row(r).Set(1, $"row-{r}");

        var data = SheetDataOf(sheet);          // hatch access — invalidates
        sheet["A1"].GetString().Should().Be("row-1"); // facade call — rebuilds the cache
        // Out-of-contract: mutate via the STORED reference, no re-acquire.
        data.Elements<S.Row>().Single(r => r.RowIndex!.Value == 3u).Remove();

        // The per-hit liveness check must detect the detached node and rebuild
        // — never serve the stale element.
        sheet["A3"].Kind.Should().Be(CellKind.Empty);
        sheet["A3"].GetString().Should().BeEmpty();
        sheet["A2"].GetString().Should().Be("row-2");
    }

    [Fact]
    public void Stored_Reference_Renumber_Is_Caught_By_The_Liveness_Check()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        sheet.Row(4).HeightInPoints = 28f;

        var data = SheetDataOf(sheet);          // hatch access — invalidates
        sheet.Row(1).Set(1, "warm");            // facade call — rebuilds the cache
        // Out-of-contract renumber via the stored reference.
        data.Elements<S.Row>().Single(r => r.RowIndex!.Value == 4u).RowIndex = 6u;

        sheet.Row(4).HeightInPoints.Should().Be(15f); // stale entry rejected
    }

    [Fact]
    public void Stored_Reference_Tail_Append_Is_Caught_By_The_Tail_Check()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        for (int r = 1; r <= 3; r++) sheet.Row(r).Set(1, r);

        var data = SheetDataOf(sheet);          // hatch access — invalidates
        sheet["A1"].GetNumber().Should().Be(1); // facade call — rebuilds the cache
        // Out-of-contract append via the stored reference.
        data.AppendChild(new S.Row { RowIndex = 9u });

        // AppendRow must land AFTER the out-of-band row 9, not collide at 4.
        var appended = sheet.AppendRow();
        appended.Index.Should().Be(10);
        appended.Set(1, "tail");
        sheet["A10"].GetString().Should().Be("tail");
    }

    // ---- The cache survives engine operations that move cells --------------

    [Fact]
    public void SortRange_Then_Continued_Writes_Stay_Coherent()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        sheet.Row(1).Set(1, 3).Set(2, "c");
        sheet.Row(2).Set(1, 1).Set(2, "a");
        sheet.Row(3).Set(1, 2).Set(2, "b");

        sheet.SortRange("A1:B3", SortKey.Asc(1));

        sheet["B1"].GetString().Should().Be("a");
        sheet["B2"].GetString().Should().Be("b");
        sheet["B3"].GetString().Should().Be("c");

        // Post-sort appends and reads keep using the (still valid) cache.
        sheet.AppendRow().Set(1, 4).Set(2, "d");
        sheet["B4"].GetString().Should().Be("d");
        sheet.LastRowNumber.Should().Be(4);
    }

    [Fact]
    public void Save_Reopen_Round_Trip_After_Hatch_Mutation()
    {
        var path = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(), $"netxlsx-rowcache-{System.Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("Data");
                for (int r = 1; r <= 6; r++) sheet.Row(r).Set(1, $"row-{r}");
                SheetDataOf(sheet).Elements<S.Row>()
                    .Single(r => r.RowIndex!.Value == 2u).Remove();
                sheet.AppendRow().Set(1, "appended"); // row 7
                wb.Save(path);
            }

            using var reopened = Workbook.Open(path);
            var s = reopened["Data"];
            s["A1"].GetString().Should().Be("row-1");
            s["A2"].Kind.Should().Be(CellKind.Empty);
            s["A7"].GetString().Should().Be("appended");
            s.LastRowNumber.Should().Be(7);
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }
}
