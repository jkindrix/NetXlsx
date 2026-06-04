// I-82 engine swap — structure slice (5b): row/column grouping (outline)
// conformance. Mirrors the NPOI engine's GroupRows / UngroupRows / GroupColumns /
// UngroupColumns / SetRowGroupCollapsed contract (decision I-71, all 1-based) on
// the Open XML SDK engine, reading <row @outlineLevel>/<col @outlineLevel> and
// <sheetFormatPr @outlineLevelRow/@outlineLevelCol> directly. Cross-checked against
// NetXlsx.Tests.GroupingTests (the NPOI-engine grouping contract).

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.Tests.Engine;

public class GroupingTests
{
    private static S.Worksheet Ws(IWorkbook wb)
        => wb.Underlying.WorkbookPart!.WorksheetParts.Single().Worksheet!;

    private static byte RowLevel(IWorkbook wb, int row1Based)
    {
        var row = Ws(wb).GetFirstChild<S.SheetData>()!.Elements<S.Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == (uint)row1Based);
        return row?.OutlineLevel?.Value ?? 0;
    }

    private static byte ColLevel(IWorkbook wb, int col1Based)
    {
        var cols = Ws(wb).GetFirstChild<S.Columns>();
        var col = cols?.Elements<S.Column>()
            .FirstOrDefault(c => c.Min?.Value <= (uint)col1Based && (uint)col1Based <= c.Max?.Value);
        return col?.OutlineLevel?.Value ?? 0;
    }

    private static bool RowHidden(IWorkbook wb, int row1Based)
    {
        var row = Ws(wb).GetFirstChild<S.SheetData>()!.Elements<S.Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == (uint)row1Based);
        return row?.Hidden?.Value ?? false;
    }

    [Fact]
    public void GroupRows_Sets_OutlineLevel_One()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.Row(1); s.Row(2); s.Row(3);
        s.GroupRows(1, 3);

        RowLevel(wb, 1).Should().Be(1);
        RowLevel(wb, 2).Should().Be(1);
        RowLevel(wb, 3).Should().Be(1);
        Ws(wb).GetFirstChild<S.SheetFormatProperties>()!.OutlineLevelRow!.Value.Should().Be(1);
    }

    [Fact]
    public void GroupRows_Nested_Increments_Level()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        for (int i = 1; i <= 5; i++) s.Row(i);
        s.GroupRows(1, 5);
        s.GroupRows(2, 4);

        RowLevel(wb, 1).Should().Be(1);
        RowLevel(wb, 2).Should().Be(2);
        RowLevel(wb, 3).Should().Be(2);
        RowLevel(wb, 4).Should().Be(2);
        RowLevel(wb, 5).Should().Be(1);
        Ws(wb).GetFirstChild<S.SheetFormatProperties>()!.OutlineLevelRow!.Value.Should().Be(2);
    }

    [Fact]
    public void UngroupRows_Decrements_Level()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        for (int i = 1; i <= 3; i++) s.Row(i).Set(1, $"r{i}");
        s.GroupRows(1, 3);
        s.UngroupRows(1, 3);

        RowLevel(wb, 1).Should().Be(0);
        // The deepest level is now 0 — the tracking attribute is cleared.
        Ws(wb).GetFirstChild<S.SheetFormatProperties>()?.OutlineLevelRow.Should().BeNull();
    }

    [Fact]
    public void GroupColumns_Sets_OutlineLevel()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.GroupColumns(1, 3);

        ColLevel(wb, 1).Should().Be(1);
        ColLevel(wb, 2).Should().Be(1);
        ColLevel(wb, 3).Should().Be(1);
        Ws(wb).GetFirstChild<S.SheetFormatProperties>()!.OutlineLevelColumn!.Value.Should().Be(1);
    }

    [Fact]
    public void UngroupColumns_Removes_Outline()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.GroupColumns(1, 3);
        s.UngroupColumns(1, 3);

        ColLevel(wb, 1).Should().Be(0);
        ColLevel(wb, 2).Should().Be(0);
        ColLevel(wb, 3).Should().Be(0);
    }

    [Fact]
    public void SetRowGroupCollapsed_Hides_Grouped_Rows()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        for (int i = 1; i <= 4; i++) s.Row(i);
        s.GroupRows(2, 4);
        s.SetRowGroupCollapsed(2, true);

        RowHidden(wb, 2).Should().BeTrue();
        RowHidden(wb, 3).Should().BeTrue();
        RowHidden(wb, 4).Should().BeTrue();
    }

    [Fact]
    public void SetRowGroupCollapsed_Then_Expand_Unhides()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        for (int i = 1; i <= 4; i++) s.Row(i).Set(1, $"r{i}");
        s.GroupRows(2, 4);
        s.SetRowGroupCollapsed(2, true);
        s.SetRowGroupCollapsed(2, false);

        RowHidden(wb, 2).Should().BeFalse();
        RowHidden(wb, 3).Should().BeFalse();
        RowHidden(wb, 4).Should().BeFalse();
    }

    [Fact]
    public void GroupRows_Rejects_Zero_And_Reversed()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        ((Action)(() => s.GroupRows(0, 3))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => s.GroupRows(5, 3))).Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void GroupColumns_Rejects_Zero()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        ((Action)(() => s.GroupColumns(0, 3))).Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void SetRowGroupCollapsed_Rejects_Zero()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        ((Action)(() => s.SetRowGroupCollapsed(0, true))).Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void GroupRows_RoundTrips_Through_Save_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-group-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddSheet("S");
                for (int i = 1; i <= 5; i++) s.Row(i).Set(1, $"Row {i}");
                s.GroupRows(2, 4);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                RowLevel(wb, 2).Should().Be(1);
                RowLevel(wb, 3).Should().Be(1);
                RowLevel(wb, 4).Should().Be(1);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Grouping_Is_Schema_Valid()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        for (int i = 1; i <= 6; i++) s.Row(i).Set(1, $"r{i}");
        s.GroupRows(2, 5);
        s.GroupRows(3, 4);
        s.GroupColumns(2, 4);
        s.SetRowGroupCollapsed(3, true);
        OpenXmlValidationGate.AssertValid(wb);
    }
}
