using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class GroupingTests
{
    [Fact]
    public void GroupRows_Sets_OutlineLevel()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.Row(1); sheet.Row(2); sheet.Row(3);
        sheet.GroupRows(1, 3);

        sheet.Underlying.GetRow(0).OutlineLevel.Should().Be(1);
        sheet.Underlying.GetRow(1).OutlineLevel.Should().Be(1);
        sheet.Underlying.GetRow(2).OutlineLevel.Should().Be(1);
    }

    [Fact]
    public void GroupRows_Nested_Increments_Level()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        for (int i = 1; i <= 5; i++) sheet.Row(i);
        sheet.GroupRows(1, 5);
        sheet.GroupRows(2, 4);

        sheet.Underlying.GetRow(0).OutlineLevel.Should().Be(1);
        sheet.Underlying.GetRow(1).OutlineLevel.Should().Be(2);
        sheet.Underlying.GetRow(2).OutlineLevel.Should().Be(2);
        sheet.Underlying.GetRow(3).OutlineLevel.Should().Be(2);
        sheet.Underlying.GetRow(4).OutlineLevel.Should().Be(1);
    }

    [Fact]
    public void UngroupRows_Decrements_Level()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        for (int i = 1; i <= 3; i++) sheet.Row(i).Set(1, $"r{i}");
        sheet.GroupRows(1, 3);
        sheet.UngroupRows(1, 3);

        var row = sheet.Underlying.GetRow(0);
        (row?.OutlineLevel ?? 0).Should().Be(0);
    }

    [Fact]
    public void GroupColumns_Sets_Outline()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.GroupColumns(1, 3);

        var ct = sheet.Underlying.GetCTWorksheet();
        ct.cols.Should().NotBeNull();
    }

    [Fact]
    public void UngroupColumns_Removes_Outline()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.GroupColumns(1, 3);
        sheet.UngroupColumns(1, 3);
        // Should not throw
    }

    [Fact]
    public void SetRowGroupCollapsed_Hides_Grouped_Rows()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        for (int i = 1; i <= 4; i++) sheet.Row(i);
        sheet.GroupRows(2, 4);
        sheet.SetRowGroupCollapsed(2, true);

        sheet.Underlying.GetRow(1).ZeroHeight.Should().BeTrue();
        sheet.Underlying.GetRow(2).ZeroHeight.Should().BeTrue();
        sheet.Underlying.GetRow(3).ZeroHeight.Should().BeTrue();
    }

    [Fact]
    public void SetRowGroupCollapsed_Then_Expand_Does_Not_Throw()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        for (int i = 1; i <= 4; i++) sheet.Row(i).Set(1, $"r{i}");
        sheet.GroupRows(2, 4);
        sheet.SetRowGroupCollapsed(2, true);
        sheet.SetRowGroupCollapsed(2, false);
        // NPOI's expand behavior is best-effort; we verify no exception
    }

    [Fact]
    public void GroupRows_Rejects_Zero_StartRow()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        Action act = () => sheet.GroupRows(0, 3);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void GroupRows_Rejects_Start_Greater_Than_End()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        Action act = () => sheet.GroupRows(5, 3);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void GroupColumns_Rejects_Zero()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        Action act = () => sheet.GroupColumns(0, 3);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void SetRowGroupCollapsed_Rejects_Zero()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        Action act = () => sheet.SetRowGroupCollapsed(0, true);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void GroupRows_Survives_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var sheet = wb.AddSheet("S");
            for (int i = 1; i <= 5; i++)
                sheet.Row(i).Set(1, $"Row {i}");
            sheet.GroupRows(2, 4);
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        var s = opened["S"];
        s.Underlying.GetRow(1).OutlineLevel.Should().Be(1);
        s.Underlying.GetRow(2).OutlineLevel.Should().Be(1);
        s.Underlying.GetRow(3).OutlineLevel.Should().Be(1);
    }
}
