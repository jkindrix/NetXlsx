// I-82 engine swap — structure slice: merged-region conformance.
//
// Mirrors the NPOI engine's MergeCells / MergeCellsStyled / UnmergeCells /
// MergedRanges contract on the Open XML SDK engine: canonical A1:C3 round-trip,
// the 1x1 no-op (I-38), overlap rejection (§6.4), exact-range unmerge with a
// silent miss, the empty-container drop (CT_MergeCells requires >= 1 child), and
// the lesson #4 style-all-cells behavior of MergeCellsStyled. Save/Open
// round-trips the regions through real OOXML.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml.Packaging;
using NetXlsx;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.OoxmlEngine.Tests;

public class MergeTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-merge-{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void MergeCells_Adds_A_Canonical_Region()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.MergeCells("A1:C3");
        s.MergedRanges.Should().ContainSingle().Which.Should().Be("A1:C3");
    }

    [Fact]
    public void MergeCells_Normalizes_A_Reversed_Range()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.MergeCells("C3:A1");
        s.MergedRanges.Should().ContainSingle().Which.Should().Be("A1:C3");
    }

    [Fact]
    public void MergeCells_1x1_Is_A_NoOp()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.MergeCells("B2:B2");
        s.MergedRanges.Should().BeEmpty();
        // No <mergeCells> element should have been created.
        var ws = wb.Underlying.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        ws.GetFirstChild<S.MergeCells>().Should().BeNull();
    }

    [Fact]
    public void Multiple_NonOverlapping_Merges_Coexist()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.MergeCells("A1:B1");
        s.MergeCells("A2:B2");
        s.MergeCells("D1:D5");
        s.MergedRanges.Should().Equal("A1:B1", "A2:B2", "D1:D5"); // insertion order
    }

    [Fact]
    public void Overlapping_Merge_Throws()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.MergeCells("A1:C3");

        var act = () => s.MergeCells("B2:D4");
        act.Should().Throw<InvalidOperationException>()
            .WithMessage("*overlaps existing merged region A1:C3*");
        // The failed merge left the region set unchanged.
        s.MergedRanges.Should().ContainSingle().Which.Should().Be("A1:C3");
    }

    [Fact]
    public void UnmergeCells_Removes_The_Exact_Region()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.MergeCells("A1:B2");
        s.MergeCells("D1:E2");
        s.UnmergeCells("A1:B2");
        s.MergedRanges.Should().ContainSingle().Which.Should().Be("D1:E2");
    }

    [Fact]
    public void UnmergeCells_Of_A_NonMatching_Range_Is_A_Silent_NoOp()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.MergeCells("A1:C3");
        s.UnmergeCells("A1:B2"); // sub-range, not an exact match
        s.UnmergeCells("Z9:Z10"); // unrelated
        s.MergedRanges.Should().ContainSingle().Which.Should().Be("A1:C3");
    }

    [Fact]
    public void Unmerging_The_Last_Region_Drops_The_Container()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.MergeCells("A1:C3");
        s.UnmergeCells("A1:C3");
        s.MergedRanges.Should().BeEmpty();

        // CT_MergeCells requires >= 1 child — a childless <mergeCells/> would be
        // schema-invalid, so the container must be gone, not just empty.
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void MergeCellsStyled_Styles_Every_Cell_In_The_Range_And_Merges()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.MergeCellsStyled("A1:C1", new CellStyle { Bold = true, Background = Color.FromRgb(0xFF, 0xFF, 0x00) });

        s.MergedRanges.Should().ContainSingle().Which.Should().Be("A1:C1");
        // Lesson #4: borders render from boundary cells, so every cell — not just
        // the A1 anchor — must carry the style. Check the C1 boundary cell.
        s["A1"].GetStyle().Bold.Should().BeTrue();
        s["C1"].GetStyle().Bold.Should().BeTrue();
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Range_Merge_Delegates_To_The_Sheet_Surface()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        var range = s.Range("B2:D4").Merge();

        range.Should().NotBeNull();
        s.MergedRanges.Should().ContainSingle().Which.Should().Be("B2:D4");
        // Same contract as MergeCells: an overlapping merge throws.
        Action act = () => s.Range("C3:E5").Merge();
        act.Should().Throw<InvalidOperationException>();
    }

    [Fact]
    public void Merges_RoundTrip_Through_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var s = wb.AddSheet("S");
                s.MergeCells("A1:C3");
                s.MergeCells("E1:E10");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                wb["S"].MergedRanges.Should().Equal("A1:C3", "E1:E10");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
