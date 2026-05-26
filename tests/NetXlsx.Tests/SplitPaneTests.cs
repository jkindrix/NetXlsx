using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class SplitPaneTests
{
    [Fact]
    public void CreateSplitPane_Sets_Split_On_Sheet()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.CreateSplitPane(2000, 3000);

        var pane = sheet.Underlying.PaneInformation;
        pane.Should().NotBeNull();
        pane!.IsFreezePane().Should().BeFalse();
        pane.VerticalSplitPosition.Should().Be(2000);
        pane.HorizontalSplitPosition.Should().Be(3000);
    }

    [Fact]
    public void CreateSplitPane_Horizontal_Only()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.CreateSplitPane(3000, 0);

        var pane = sheet.Underlying.PaneInformation;
        pane.Should().NotBeNull();
        pane!.VerticalSplitPosition.Should().Be(3000);
        pane.HorizontalSplitPosition.Should().Be(0);
    }

    [Fact]
    public void CreateSplitPane_Vertical_Only()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.CreateSplitPane(0, 4000);

        var pane = sheet.Underlying.PaneInformation;
        pane.Should().NotBeNull();
        pane!.HorizontalSplitPosition.Should().Be(4000);
    }

    [Fact]
    public void CreateSplitPane_Rejects_Negative_X()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        Action act = () => sheet.CreateSplitPane(-1, 0);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void CreateSplitPane_Rejects_Negative_Y()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        Action act = () => sheet.CreateSplitPane(0, -1);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void CreateSplitPane_Survives_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            wb.AddSheet("S").CreateSplitPane(2500, 3500);
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        var pane = opened["S"].Underlying.PaneInformation;
        pane.Should().NotBeNull();
        pane!.IsFreezePane().Should().BeFalse();
        pane.VerticalSplitPosition.Should().Be(2500);
        pane.HorizontalSplitPosition.Should().Be(3500);
    }

    [Fact]
    public void CreateSplitPane_Replaces_Prior_Freeze()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.FreezePane(2, 1);
        sheet.CreateSplitPane(2000, 3000);

        var pane = sheet.Underlying.PaneInformation;
        pane.Should().NotBeNull();
        pane!.IsFreezePane().Should().BeFalse();
    }
}
