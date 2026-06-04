// Split-pane behavior. Pane state has no public read-back, so tests
// assert on the persisted <sheetView>/<pane> element via SavedOoxml —
// engine-agnostic, no .Underlying reach-through (I-82 cutover phase 1).
// Attribute mapping: @xSplit is the vertical split line (x position),
// @ySplit the horizontal one; @state defaults to "split", "frozen" marks
// a freeze pane.

using System;
using System.IO;
using System.Xml.Linq;
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

        var pane = Pane(wb);
        pane.Should().NotBeNull();
        IsFrozen(pane!).Should().BeFalse();
        XSplit(pane!).Should().Be(2000);
        YSplit(pane!).Should().Be(3000);
    }

    [Fact]
    public void CreateSplitPane_Horizontal_Only()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.CreateSplitPane(3000, 0);

        var pane = Pane(wb);
        pane.Should().NotBeNull();
        XSplit(pane!).Should().Be(3000);
        YSplit(pane!).Should().Be(0);
    }

    [Fact]
    public void CreateSplitPane_Vertical_Only()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.CreateSplitPane(0, 4000);

        var pane = Pane(wb);
        pane.Should().NotBeNull();
        YSplit(pane!).Should().Be(4000);
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
        var pane = Pane(opened);
        pane.Should().NotBeNull();
        IsFrozen(pane!).Should().BeFalse();
        XSplit(pane!).Should().Be(2500);
        YSplit(pane!).Should().Be(3500);
    }

    [Fact]
    public void CreateSplitPane_Replaces_Prior_Freeze()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.FreezePane(2, 1);
        sheet.CreateSplitPane(2000, 3000);

        var pane = Pane(wb);
        pane.Should().NotBeNull();
        IsFrozen(pane!).Should().BeFalse();
    }

    // ---- helpers ------------------------------------------------------

    internal static XElement? Pane(IWorkbook wb)
        => SavedOoxml.SheetXml(wb).Root!
            .Element(SavedOoxml.Main + "sheetViews")!
            .Element(SavedOoxml.Main + "sheetView")!
            .Element(SavedOoxml.Main + "pane");

    internal static double XSplit(XElement pane)
        => (double?)pane.Attribute("xSplit") ?? 0;

    internal static double YSplit(XElement pane)
        => (double?)pane.Attribute("ySplit") ?? 0;

    /// <summary>pane/@state; OOXML default is "split".</summary>
    internal static bool IsFrozen(XElement pane)
        => (string?)pane.Attribute("state") == "frozen";
}
