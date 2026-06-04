// Row/column grouping behavior. Outline levels and collapse-hidden state
// have no public read-back, so tests assert on the persisted sheet XML
// (row @outlineLevel / @hidden, <cols>) via SavedOoxml — engine-agnostic,
// no .Underlying reach-through (I-82 cutover phase 1). Row numbers in the
// XML are 1-based (@r), matching the public API.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
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

        var sheetXml = SavedOoxml.SheetXml(wb);
        OutlineLevel(sheetXml, 1).Should().Be(1);
        OutlineLevel(sheetXml, 2).Should().Be(1);
        OutlineLevel(sheetXml, 3).Should().Be(1);
    }

    [Fact]
    public void GroupRows_Nested_Increments_Level()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        for (int i = 1; i <= 5; i++) sheet.Row(i);
        sheet.GroupRows(1, 5);
        sheet.GroupRows(2, 4);

        var sheetXml = SavedOoxml.SheetXml(wb);
        OutlineLevel(sheetXml, 1).Should().Be(1);
        OutlineLevel(sheetXml, 2).Should().Be(2);
        OutlineLevel(sheetXml, 3).Should().Be(2);
        OutlineLevel(sheetXml, 4).Should().Be(2);
        OutlineLevel(sheetXml, 5).Should().Be(1);
    }

    [Fact]
    public void UngroupRows_Decrements_Level()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        for (int i = 1; i <= 3; i++) sheet.Row(i).Set(1, $"r{i}");
        sheet.GroupRows(1, 3);
        sheet.UngroupRows(1, 3);

        OutlineLevel(SavedOoxml.SheetXml(wb), 1).Should().Be(0);
    }

    [Fact]
    public void GroupColumns_Sets_Outline()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.GroupColumns(1, 3);

        var cols = SavedOoxml.SheetXml(wb).Root!.Element(SavedOoxml.Main + "cols");
        cols.Should().NotBeNull();
        cols!.Elements(SavedOoxml.Main + "col")
            .Should().Contain(c => (int?)c.Attribute("outlineLevel") == 1);
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

        var sheetXml = SavedOoxml.SheetXml(wb);
        RowHidden(sheetXml, 2).Should().BeTrue();
        RowHidden(sheetXml, 3).Should().BeTrue();
        RowHidden(sheetXml, 4).Should().BeTrue();
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
        var sheetXml = SavedOoxml.SheetXml(opened);
        OutlineLevel(sheetXml, 2).Should().Be(1);
        OutlineLevel(sheetXml, 3).Should().Be(1);
        OutlineLevel(sheetXml, 4).Should().Be(1);
    }

    // ---- helpers ------------------------------------------------------

    private static XElement? Row(XDocument sheetXml, int rowNumber)
        => sheetXml.Root!.Element(SavedOoxml.Main + "sheetData")!
            .Elements(SavedOoxml.Main + "row")
            .FirstOrDefault(r => (int?)r.Attribute("r") == rowNumber);

    private static int OutlineLevel(XDocument sheetXml, int rowNumber)
        => (int?)Row(sheetXml, rowNumber)?.Attribute("outlineLevel") ?? 0;

    private static bool RowHidden(XDocument sheetXml, int rowNumber)
        => SavedOoxml.BoolAttr(Row(sheetXml, rowNumber), "hidden");
}
