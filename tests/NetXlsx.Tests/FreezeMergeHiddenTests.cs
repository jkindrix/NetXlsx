// Coverage for sub-slice A: freeze panes, merge cells, hidden rows/sheets,
// gridlines toggle. Each landed as a property/method on ISheet or IRow
// per design §6.3 / §6.6 / §6.4.

using System;
using System.Collections.Generic;
using System.IO;
using FluentAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class FreezeMergeHiddenTests
{
    private static readonly string[] s_twoMergedRanges = { "A1:C1", "A2:B5" };

    // ---- Freeze panes --------------------------------------------------

    [Fact]
    public void FreezeRows_Sets_Top_Freeze_Pane()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.AppendRow().Set(1, "Header");
        sheet.FreezeRows(1);

        var pane = sheet.Underlying.PaneInformation;
        pane.Should().NotBeNull();
        pane.HorizontalSplitPosition.Should().Be(1, "row 1 is frozen above the split");
        pane.VerticalSplitPosition.Should().Be(0);
    }

    [Fact]
    public void FreezeColumns_Sets_Left_Freeze_Pane()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.FreezeColumns(2);

        var pane = sheet.Underlying.PaneInformation;
        pane.VerticalSplitPosition.Should().Be(2);
        pane.HorizontalSplitPosition.Should().Be(0);
    }

    [Fact]
    public void FreezePane_Sets_Both_Splits()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.FreezePane(1, 3);

        var pane = sheet.Underlying.PaneInformation;
        pane.HorizontalSplitPosition.Should().Be(1);
        pane.VerticalSplitPosition.Should().Be(3);
    }

    [Fact]
    public void FreezePane_Negative_Args_Throw()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet.FreezePane(-1, 0))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => sheet.FreezePane(0, -1))).Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void FreezePane_Survives_Save_Open_Round_Trip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"freeze-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("S");
                sheet.AppendRow().Set(1, "Header");
                sheet.FreezeRows(1);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var pane = wb["S"].Underlying.PaneInformation;
                pane.HorizontalSplitPosition.Should().Be(1);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Merge cells ---------------------------------------------------

    [Fact]
    public void MergeCells_Adds_The_Range()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.MergeCells("A1:C3");
        sheet.MergedRanges.Should().ContainSingle().Which.Should().Be("A1:C3");
    }

    [Fact]
    public void MergeCells_1x1_Range_Is_NoOp_Per_Decision_I38()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.MergeCells("A1");
        sheet.MergedRanges.Should().BeEmpty();
    }

    [Fact]
    public void MergeCells_Overlapping_Existing_Throws_InvalidOperation()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.MergeCells("A1:C3");

        ((Action)(() => sheet.MergeCells("B2:D4"))).Should()
            .Throw<InvalidOperationException>()
            .WithMessage("*overlaps existing*A1:C3*");
    }

    [Fact]
    public void MergeCells_Adjacent_Non_Overlapping_Range_Succeeds()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.MergeCells("A1:C3");
        sheet.MergeCells("D1:F3");   // shares an edge but does not overlap
        sheet.MergedRanges.Should().HaveCount(2);
    }

    [Fact]
    public void UnmergeCells_Removes_Exact_Match()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.MergeCells("A1:C3");
        sheet.UnmergeCells("A1:C3");
        sheet.MergedRanges.Should().BeEmpty();
    }

    [Fact]
    public void UnmergeCells_Non_Matching_Range_Is_Silent_NoOp_Per_Design_6_4()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.MergeCells("A1:C3");
        sheet.UnmergeCells("A1:B2");   // partial — not an exact match
        sheet.MergedRanges.Should().ContainSingle().Which.Should().Be("A1:C3");
    }

    [Fact]
    public void MergeCells_Survives_Save_Open_Round_Trip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"merge-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("S");
                sheet.MergeCells("A1:C1");
                sheet.MergeCells("A2:B5");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["S"].MergedRanges.Should().BeEquivalentTo(s_twoMergedRanges);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void MergeCells_Bad_Range_Throws_InvalidCellAddress()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet.MergeCells("not-a-range"))).Should()
            .Throw<InvalidCellAddressException>();
    }

    // ---- Hidden sheet --------------------------------------------------

    [Fact]
    public void Sheet_Hidden_Defaults_False()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S").Hidden.Should().BeFalse();
    }

    [Fact]
    public void Sheet_Hidden_Set_True_Then_False_Round_Trips()
    {
        var path = Path.Combine(Path.GetTempPath(), $"hidden-sheet-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                // First sheet must remain visible — Excel requires at least
                // one visible sheet. Add a second to hide.
                wb.AddSheet("Visible");
                wb.AddSheet("HideMe").Hidden = true;
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["Visible"].Hidden.Should().BeFalse();
                wb["HideMe"].Hidden.Should().BeTrue();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- ShowGridlines -------------------------------------------------

    [Fact]
    public void ShowGridlines_Toggles_And_Round_Trips()
    {
        var path = Path.Combine(Path.GetTempPath(), $"gridlines-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("S");
                sheet.ShowGridlines.Should().BeTrue("Excel defaults to showing gridlines");
                sheet.ShowGridlines = false;
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["S"].ShowGridlines.Should().BeFalse();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Hidden row ----------------------------------------------------

    [Fact]
    public void Row_Hidden_Defaults_False_And_Toggles()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var row = sheet.AppendRow().Set(1, "data");
        row.Hidden.Should().BeFalse();
        row.Hidden = true;
        row.Hidden.Should().BeTrue();
    }

    [Fact]
    public void Row_Hidden_Survives_Save_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"hidden-row-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("S");
                sheet.AppendRow().Set(1, "visible");
                var secondRow = sheet.AppendRow().Set(1, "hidden");
                secondRow.Hidden = true;
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var sheet = wb["S"];
                sheet.Row(1).Hidden.Should().BeFalse();
                sheet.Row(2).Hidden.Should().BeTrue();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
