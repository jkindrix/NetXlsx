// I-82 engine swap — structure slice (5b): frozen/split pane conformance.
//
// Mirrors the NPOI engine's FreezeRows / FreezeColumns / FreezePane (column-first
// CreateFreezePane mapping) / CreateSplitPane contract on the Open XML SDK engine,
// reading the <sheetView><pane> node directly: xSplit = frozen columns,
// ySplit = frozen rows, topLeftCell = first unfrozen cell, activePane + a matching
// <selection>, state = frozen (panes) or split (draggable). Cross-checked against
// NetXlsx.Tests.FreezeMergeHiddenTests (the NPOI-engine pane contract).

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.Tests.Engine;

public class PaneTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-pane-{Guid.NewGuid():N}.xlsx");

    private static S.Pane? PaneOf(IWorkbook wb)
        => wb.Underlying.WorkbookPart!.WorksheetParts.Single().Worksheet!
            .GetFirstChild<S.SheetViews>()?.GetFirstChild<S.SheetView>()?.GetFirstChild<S.Pane>();

    [Fact]
    public void FreezeRows_Freezes_Top_Rows()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.FreezeRows(1);

        var pane = PaneOf(wb);
        pane.Should().NotBeNull();
        pane!.VerticalSplit?.Value.Should().Be(1, "ySplit counts frozen rows");
        pane.HorizontalSplit.Should().BeNull("no columns are frozen");
        pane.State!.Value.Should().Be(S.PaneStateValues.Frozen);
        pane.TopLeftCell!.Value.Should().Be("A2");
        pane.ActivePane!.Value.Should().Be(S.PaneValues.BottomLeft);
    }

    [Fact]
    public void FreezeColumns_Freezes_Left_Columns()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.FreezeColumns(2);

        var pane = PaneOf(wb);
        pane.Should().NotBeNull();
        pane!.HorizontalSplit?.Value.Should().Be(2, "xSplit counts frozen columns");
        pane.VerticalSplit.Should().BeNull("no rows are frozen");
        pane.State!.Value.Should().Be(S.PaneStateValues.Frozen);
        pane.TopLeftCell!.Value.Should().Be("C1");
        pane.ActivePane!.Value.Should().Be(S.PaneValues.TopRight);
    }

    [Fact]
    public void FreezePane_Freezes_Both_Splits()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.FreezePane(1, 3); // 1 row, 3 columns

        var pane = PaneOf(wb);
        pane.Should().NotBeNull();
        pane!.HorizontalSplit?.Value.Should().Be(3);
        pane.VerticalSplit?.Value.Should().Be(1);
        pane.State!.Value.Should().Be(S.PaneStateValues.Frozen);
        pane.TopLeftCell!.Value.Should().Be("D2", "first unfrozen cell is below 1 row and right of 3 columns");
        pane.ActivePane!.Value.Should().Be(S.PaneValues.BottomRight);
    }

    [Fact]
    public void FreezePane_Emits_A_Selection_In_The_Active_Pane()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S").FreezePane(1, 1);

        var view = wb.Underlying.WorkbookPart!.WorksheetParts.Single().Worksheet!
            .GetFirstChild<S.SheetViews>()!.GetFirstChild<S.SheetView>()!;
        var sel = view.Elements<S.Selection>().Single();
        sel.Pane!.Value.Should().Be(S.PaneValues.BottomRight);
        // <pane> must precede <selection> in CT_SheetView.
        view.GetFirstChild<S.Pane>().Should().NotBeNull();
        view.ChildElements.ToList().IndexOf(view.GetFirstChild<S.Pane>()!)
            .Should().BeLessThan(view.ChildElements.ToList().IndexOf(sel));
    }

    [Fact]
    public void FreezePane_Negative_Args_Throw()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        ((Action)(() => s.FreezePane(-1, 0))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => s.FreezePane(0, -1))).Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void FreezePane_Zero_Zero_Clears_An_Existing_Freeze()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.FreezePane(2, 2);
        PaneOf(wb).Should().NotBeNull();
        s.FreezePane(0, 0);
        PaneOf(wb).Should().BeNull("a (0,0) freeze removes the <pane>");
    }

    [Fact]
    public void CreateSplitPane_Uses_Split_State()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.CreateSplitPane(2000, 1000);

        var pane = PaneOf(wb);
        pane.Should().NotBeNull();
        pane!.HorizontalSplit?.Value.Should().Be(2000);
        pane.VerticalSplit?.Value.Should().Be(1000);
        pane.State!.Value.Should().Be(S.PaneStateValues.Split);
        pane.ActivePane!.Value.Should().Be(S.PaneValues.BottomRight);
        pane.TopLeftCell!.Value.Should().Be("A1");
    }

    [Fact]
    public void CreateSplitPane_Negative_Args_Throw()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        ((Action)(() => s.CreateSplitPane(-1, 0))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => s.CreateSplitPane(0, -1))).Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void FreezePane_RoundTrips_Through_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S").FreezePane(1, 2);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var pane = PaneOf(wb);
                pane.Should().NotBeNull();
                pane!.HorizontalSplit?.Value.Should().Be(2);
                pane.VerticalSplit?.Value.Should().Be(1);
                pane.State!.Value.Should().Be(S.PaneStateValues.Frozen);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Panes_Are_Schema_Valid()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Freeze").FreezePane(2, 3);
        wb.AddSheet("Split").CreateSplitPane(1500, 1500);
        OpenXmlValidationGate.AssertValid(wb);
    }
}
