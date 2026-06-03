// I-82 engine swap — charts slice (the I-75 surface): chart conformance.
//
// Mirrors the NPOI engine's ChartTests contract on the Open XML SDK engine
// (decision I-75): ISheet.AddChart for all six ChartTypes, the optional title,
// IChart.SetTitle, argument validation, and a file round-trip. It additionally
// pins the SDK engine's own emission (oracle-dumped from the NPOI engine per
// SDK-quirk #11 before implementing):
//   - twoCellAnchor with the EXCLUSIVE end cell (the picture/shape convention,
//     quirk #10) holding an xdr:graphicFrame whose r:id resolves to a ChartPart;
//   - strCache/numCache snapshots: ptCount = the full range size, only
//     type-matching cells get a <c:pt> (string cells / numeric+date cells);
//   - sheet-qualified absolute formula refs, quoted when the name needs it,
//     single-cell ranges collapsed ("Data!$A$1");
//   - the documented conformance-positive divergences from NPOI (quirk #14):
//     pie dPt list emitted in CT_PieSer schema order (before cat), no dangling
//     pie axes, scatter plotted on two value axes, nonzero cNvPr ids;
//   - the NPOI escape hatch (IChart.Underlying) throws NotSupportedException.
// Emission parity vs the NPOI engine is asserted by the projection test at the
// bottom (charts have no public read-back beyond IChart itself), normalizing
// the divergences above.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml.Packaging;
using NetXlsx;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace NetXlsx.OoxmlEngine.Tests;

public class ChartTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-chart-{Guid.NewGuid():N}.xlsx");

    private static ISheet CreateSheetWithData(IWorkbook wb)
    {
        var s = wb.AddSheet("Data");
        s["A1"].SetString("Jan"); s["B1"].SetNumber(100);
        s["A2"].SetString("Feb"); s["B2"].SetNumber(150);
        s["A3"].SetString("Mar"); s["B3"].SetNumber(200);
        s["A4"].SetString("Apr"); s["B4"].SetNumber(120);
        return s;
    }

    private static ChartPart SingleChartPartOf(IWorkbook wb)
        => wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single()
            .GetPartsOfType<DrawingsPart>().Single().ChartParts.Single();

    private static XDR.TwoCellAnchor SingleAnchorOf(IWorkbook wb)
        => wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single()
            .GetPartsOfType<DrawingsPart>().Single().WorksheetDrawing!
            .Elements<XDR.TwoCellAnchor>().Single();

    // ---- creation, one per type ---------------------------------------------

    [Theory]
    [InlineData(ChartType.Line, "lineChart")]
    [InlineData(ChartType.Bar, "barChart")]
    [InlineData(ChartType.Column, "barChart")]
    [InlineData(ChartType.Pie, "pieChart")]
    [InlineData(ChartType.Area, "areaChart")]
    public void AddChart_Creates_Chart_With_Right_Plot_Element(ChartType type, string plotElement)
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        var chart = s.AddChart(type, "D1", "K15", "A1:A4", "B1:B4");

        chart.Should().NotBeNull();
        chart.Type.Should().Be(type);
        chart.Sheet.Should().BeSameAs(s);
        var plotArea = SingleChartPartOf(wb).ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        plotArea.ChildElements.Should().Contain(e => e.LocalName == plotElement);
    }

    [Fact]
    public void AddChart_Scatter_Creates_Chart()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("Data");
        s["A1"].SetNumber(1); s["B1"].SetNumber(10);
        s["A2"].SetNumber(2); s["B2"].SetNumber(20);
        s["A3"].SetNumber(3); s["B3"].SetNumber(30);

        var chart = s.AddChart(ChartType.Scatter, "D1", "K15", "A1:A3", "B1:B3");
        chart.Type.Should().Be(ChartType.Scatter);
        SingleChartPartOf(wb).ChartSpace!.Descendants<C.ScatterChart>().Should().HaveCount(1);
    }

    [Fact]
    public void Bar_And_Column_Differ_Only_In_BarDirection()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        s.AddChart(ChartType.Bar, "D1", "K15", "A1:A4", "B1:B4");
        s.AddChart(ChartType.Column, "D16", "K30", "A1:A4", "B1:B4");

        var dirs = wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single()
            .GetPartsOfType<DrawingsPart>().Single().ChartParts
            .Select(p => p.ChartSpace!.Descendants<C.BarDirection>().Single().Val!.InnerText)
            .OrderBy(v => v, StringComparer.Ordinal).ToArray();
        dirs.Should().Equal("bar", "col");
    }

    // ---- anchor geometry (oracle: NPOI emits from 3,0 to 11,15 for D1->K15) --

    [Fact]
    public void Anchor_End_Cell_Is_Exclusive_Like_Pictures()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");

        var anchor = SingleAnchorOf(wb);
        var from = anchor.FromMarker!;
        var to = anchor.ToMarker!;
        from.ColumnId!.Text.Should().Be("3");
        from.RowId!.Text.Should().Be("0");
        to.ColumnId!.Text.Should().Be("11");
        to.RowId!.Text.Should().Be("15");
    }

    [Fact]
    public void GraphicFrame_References_The_ChartPart_With_A_Nonzero_Id()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");

        var drawingsPart = wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single()
            .GetPartsOfType<DrawingsPart>().Single();
        var frame = SingleAnchorOf(wb).GetFirstChild<XDR.GraphicFrame>()!;
        // cNvPr/@id must be unique-nonzero (quirk #9); NPOI's id=0 is one of the
        // documented nonconformities the SDK engine does not reproduce.
        frame.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!.Id!.Value.Should().BeGreaterThan(0u);
        string rid = frame.Graphic!.GraphicData!.GetFirstChild<C.ChartReference>()!.Id!.Value!;
        drawingsPart.GetPartById(rid).Should().BeOfType<ChartPart>();
    }

    [Fact]
    public void Two_Charts_On_One_Sheet_Get_Distinct_Parts_And_Ids()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");
        s.AddChart(ChartType.Pie, "D16", "K30", "A1:A4", "B1:B4");

        var drawingsPart = wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single()
            .GetPartsOfType<DrawingsPart>().Single();
        drawingsPart.ChartParts.Should().HaveCount(2);
        var ids = drawingsPart.WorksheetDrawing!
            .Descendants<XDR.NonVisualDrawingProperties>()
            .Select(p => p.Id!.Value).ToArray();
        ids.Should().OnlyHaveUniqueItems();
        ids.Should().NotContain(0u);
    }

    [Fact]
    public void Chart_Shares_The_Drawing_With_An_Existing_Picture()
    {
        // 1×1 transparent PNG (the PictureTests fixture image).
        byte[] onePixelPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        s.AddPicture("M1", onePixelPng, ImageFormat.Png);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");

        var drawingsParts = wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single()
            .GetPartsOfType<DrawingsPart>().ToList();
        drawingsParts.Should().HaveCount(1, "picture and chart share one xdr:wsDr");
        drawingsParts[0].WorksheetDrawing!.ChildElements.Should().HaveCount(2);
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- data references + caches --------------------------------------------

    [Fact]
    public void Series_References_Are_Sheet_Qualified_Absolute()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");

        var ser = SingleChartPartOf(wb).ChartSpace!.Descendants<C.LineChartSeries>().Single();
        ser.Descendants<C.StringReference>().Single().Formula!.Text.Should().Be("Data!$A$1:$A$4");
        ser.Descendants<C.NumberReference>().Single().Formula!.Text.Should().Be("Data!$B$1:$B$4");
    }

    [Fact]
    public void Caches_Snapshot_The_Referenced_Cells()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");

        var ser = SingleChartPartOf(wb).ChartSpace!.Descendants<C.LineChartSeries>().Single();
        var strCache = ser.Descendants<C.StringCache>().Single();
        strCache.GetFirstChild<C.PointCount>()!.Val!.Value.Should().Be(4u);
        strCache.Elements<C.StringPoint>().Select(p => p.NumericValue!.Text)
            .Should().Equal("Jan", "Feb", "Mar", "Apr");
        var numCache = ser.Descendants<C.NumberingCache>().Single();
        numCache.Elements<C.NumericPoint>().Select(p => p.NumericValue!.Text)
            .Should().Equal("100", "150", "200", "120");
    }

    [Fact]
    public void Caches_Skip_Empty_And_Type_Mismatched_Cells()
    {
        // The NPOI oracle: ptCount stays the full range size; only type-matching
        // cells get a <c:pt> (empty, bool, and mismatched cells are skipped).
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("Data");
        s["A1"].SetString("Jan");
        // A2 left empty
        s["A3"].SetNumber(42);      // numeric cell in the STRING cat range
        s["A4"].SetString("Apr");
        s["B1"].SetNumber(100);
        // B2 left empty
        s["B3"].SetString("oops");  // string cell in the NUMERIC val range
        s["B4"].SetNumber(120);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");

        var ser = SingleChartPartOf(wb).ChartSpace!.Descendants<C.LineChartSeries>().Single();
        var strCache = ser.Descendants<C.StringCache>().Single();
        strCache.GetFirstChild<C.PointCount>()!.Val!.Value.Should().Be(4u);
        strCache.Elements<C.StringPoint>().Select(p => (p.Index!.Value, p.NumericValue!.Text))
            .Should().Equal((0u, "Jan"), (3u, "Apr"));
        var numCache = ser.Descendants<C.NumberingCache>().Single();
        numCache.GetFirstChild<C.PointCount>()!.Val!.Value.Should().Be(4u);
        numCache.Elements<C.NumericPoint>().Select(p => (p.Index!.Value, p.NumericValue!.Text))
            .Should().Equal((0u, "100"), (3u, "120"));
    }

    [Fact]
    public void Date_Cells_Land_In_The_Numeric_Cache_As_Serials()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("Data");
        s["A1"].SetString("When");
        s["B1"].SetDate(new DateTime(2026, 1, 1));
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A1", "B1:B1");

        var numCache = SingleChartPartOf(wb).ChartSpace!.Descendants<C.NumberingCache>().Single();
        numCache.Elements<C.NumericPoint>().Single().NumericValue!.Text
            .Should().Be(s["B1"].GetNumber()!.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void Single_Cell_Range_Collapses_The_Reference()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("Data");
        s["A1"].SetString("Q1"); s["B1"].SetNumber(1);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A1", "B1:B1");

        var ser = SingleChartPartOf(wb).ChartSpace!.Descendants<C.LineChartSeries>().Single();
        ser.Descendants<C.StringReference>().Single().Formula!.Text.Should().Be("Data!$A$1");
        ser.Descendants<C.NumberReference>().Single().Formula!.Text.Should().Be("Data!$B$1");
    }

    [Fact]
    public void Sheet_Name_Needing_Quotes_Is_Quoted_In_References()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("My Data");
        s["A1"].SetString("Q1"); s["B1"].SetNumber(1);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A1", "B1:B1");

        var ser = SingleChartPartOf(wb).ChartSpace!.Descendants<C.LineChartSeries>().Single();
        ser.Descendants<C.StringReference>().Single().Formula!.Text.Should().Be("'My Data'!$A$1");
    }

    // ---- documented divergences from NPOI (quirk #14) ------------------------

    [Fact]
    public void Pie_DataPoints_Precede_Cat_And_Cycle_Accents()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("Data");
        for (int i = 1; i <= 8; i++) { s[$"A{i}"].SetString("C" + i); s[$"B{i}"].SetNumber(i * 10); }
        s.AddChart(ChartType.Pie, "D1", "K15", "A1:A8", "B1:B8");

        var ser = SingleChartPartOf(wb).ChartSpace!.Descendants<C.PieChartSeries>().Single();
        // CT_PieSer schema order: every dPt sits before <c:cat> (NPOI emits the
        // list after <c:val>, which is schema-nonconformant).
        var children = ser.ChildElements.Select(e => e.LocalName).ToList();
        children.IndexOf("cat").Should().BeGreaterThan(children.LastIndexOf("dPt"));
        ser.Elements<C.DataPoint>().Should().HaveCount(8);
        ser.Elements<C.DataPoint>()
            .Select(d => d.ChartShapeProperties!.Descendants<A.SchemeColor>().Single().Val!.InnerText)
            .Should().Equal("accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
                "accent1", "accent2");
    }

    [Fact]
    public void Pie_PlotArea_Has_No_Dangling_Axes()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        s.AddChart(ChartType.Pie, "D1", "K15", "A1:A4", "B1:B4");

        var plotArea = SingleChartPartOf(wb).ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        plotArea.Elements<C.CategoryAxis>().Should().BeEmpty();
        plotArea.Elements<C.ValueAxis>().Should().BeEmpty();
    }

    [Fact]
    public void Scatter_Plots_On_Two_Value_Axes()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("Data");
        s["A1"].SetNumber(1); s["B1"].SetNumber(10);
        s["A2"].SetNumber(2); s["B2"].SetNumber(20);
        s.AddChart(ChartType.Scatter, "D1", "K15", "A1:A2", "B1:B2");

        var chartSpace = SingleChartPartOf(wb).ChartSpace!;
        var ser = chartSpace.Descendants<C.ScatterChartSeries>().Single();
        ser.GetFirstChild<C.XValues>()!.NumberReference!.Formula!.Text.Should().Be("Data!$A$1:$A$2");
        ser.GetFirstChild<C.YValues>()!.NumberReference!.Formula!.Text.Should().Be("Data!$B$1:$B$2");
        var plotArea = chartSpace.GetFirstChild<C.Chart>()!.PlotArea!;
        // An ECMA-376 scatter chart plots two value axes (NPOI pairs it with a
        // catAx — a documented divergence).
        plotArea.Elements<C.ValueAxis>().Should().HaveCount(2);
        plotArea.Elements<C.CategoryAxis>().Should().BeEmpty();
    }

    // ---- title ----------------------------------------------------------------

    [Fact]
    public void AddChart_With_Title_Writes_It()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4", title: "Sales");

        var title = SingleChartPartOf(wb).ChartSpace!.Descendants<C.Title>().Single();
        title.Descendants<A.Text>().Single().Text.Should().Be("Sales");
    }

    [Fact]
    public void AddChart_Without_Title_Writes_None()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");

        SingleChartPartOf(wb).ChartSpace!.Descendants<C.Title>().Should().BeEmpty();
    }

    [Fact]
    public void SetTitle_Replaces_The_Existing_Title()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        var chart = s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4", title: "First");
        chart.SetTitle("Second");

        var titles = SingleChartPartOf(wb).ChartSpace!.Descendants<C.Title>().ToList();
        titles.Should().HaveCount(1);
        titles[0].Descendants<A.Text>().Single().Text.Should().Be("Second");
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void SetTitle_On_An_Untitled_Chart_Adds_One()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        var chart = s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");
        chart.SetTitle("Late");

        var chartEl = SingleChartPartOf(wb).ChartSpace!.GetFirstChild<C.Chart>()!;
        // CT_Chart is a strict sequence; the title must precede the plotArea.
        chartEl.ChildElements[0].Should().BeOfType<C.Title>();
        chartEl.Descendants<A.Text>().Single().Text.Should().Be("Late");
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- validation surface ----------------------------------------------------

    [Fact]
    public void AddChart_Rejects_Null_Arguments()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        ((Action)(() => s.AddChart(ChartType.Line, null!, "K15", "A1:A4", "B1:B4")))
            .Should().Throw<ArgumentNullException>();
        ((Action)(() => s.AddChart(ChartType.Line, "D1", null!, "A1:A4", "B1:B4")))
            .Should().Throw<ArgumentNullException>();
        ((Action)(() => s.AddChart(ChartType.Line, "D1", "K15", null!, "B1:B4")))
            .Should().Throw<ArgumentNullException>();
        ((Action)(() => s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", null!)))
            .Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void SetTitle_Rejects_Null()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        var chart = s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");
        ((Action)(() => chart.SetTitle(null!))).Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Underlying_Throws_NotSupported_On_The_Sdk_Engine()
    {
        using var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        var chart = s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");
        ((Action)(() => _ = chart.Underlying)).Should().Throw<NotSupportedException>()
            .WithMessage("*OpenXmlDocument*");
    }

    [Fact]
    public void Members_Throw_After_Dispose()
    {
        var wb = Workbook.CreateOoxml();
        var s = CreateSheetWithData(wb);
        var chart = s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");
        wb.Dispose();

        ((Action)(() => s.AddChart(ChartType.Pie, "D1", "K15", "A1:A4", "B1:B4")))
            .Should().Throw<ObjectDisposedException>();
        ((Action)(() => _ = chart.Sheet)).Should().Throw<ObjectDisposedException>();
        ((Action)(() => _ = chart.Type)).Should().Throw<ObjectDisposedException>();
        ((Action)(() => chart.SetTitle("X"))).Should().Throw<ObjectDisposedException>();
    }

    // ---- round-trip --------------------------------------------------------------

    [Fact]
    public void AddChart_Survives_RoundTrip()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var s = wb.AddSheet("Data");
                s["A1"].SetString("Q1"); s["B1"].SetNumber(100);
                s["A2"].SetString("Q2"); s["B2"].SetNumber(200);
                s.AddChart(ChartType.Bar, "D1", "J12", "A1:A2", "B1:B2", "Revenue");
                wb.Save(path);
            }

            using var opened = Workbook.OpenOoxml(path);
            var chartPart = SingleChartPartOf(opened);
            chartPart.ChartSpace!.Descendants<C.BarChart>().Should().HaveCount(1);
            chartPart.ChartSpace!.Descendants<A.Text>().Single().Text.Should().Be("Revenue");
            OpenXmlValidationGate.AssertValid(opened);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Adding_A_Chart_To_An_Opened_Sheet_With_A_Chart_Stays_Valid()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var s = CreateSheetWithData(wb);
                s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");
                wb.Save(path);
            }

            using (var opened = Workbook.OpenOoxml(path))
            {
                opened["Data"].AddChart(ChartType.Area, "D16", "K30", "A1:A4", "B1:B4");
                var drawingsPart = opened.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single()
                    .GetPartsOfType<DrawingsPart>().Single();
                drawingsPart.ChartParts.Should().HaveCount(2);
                OpenXmlValidationGate.AssertValid(opened);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- cross-engine emission parity (charts have no public read-back beyond
    // IChart itself, so this stands in for the differential harness) ------------

    private sealed record ChartObs(
        string Plot,
        string? BarDir,
        (int Col, int Row) From,
        (int Col, int Row) To,
        string? Title,
        string?[] SeriesRefs,
        (uint Idx, string Text)[][] CachePoints,
        (uint Idx, string Accent)[] PieFills);

    [Fact]
    public void Chart_Emission_Agrees_Across_Engines()
    {
        // Normalized away (the documented quirk-#14 divergences + cosmetics):
        // pie dPt position (projected as an idx-sorted set), pie/scatter axes
        // (NPOI emits dangling/cat axes), cNvPr id/name, editAs, part URIs,
        // relationship-id format.
        static ChartObs[] EmitAndProject(Func<IWorkbook> create)
        {
            var path = Path.Combine(Path.GetTempPath(), $"netxlsx-chart-par-{Guid.NewGuid():N}.xlsx");
            try
            {
                using (var wb = create())
                {
                    foreach (ChartType type in Enum.GetValues<ChartType>())
                    {
                        var s = wb.AddSheet("D" + type);
                        if (type == ChartType.Scatter)
                        {
                            s["A1"].SetNumber(1); s["B1"].SetNumber(10);
                            s["A2"].SetNumber(2); s["B2"].SetNumber(20);
                        }
                        else
                        {
                            s["A1"].SetString("Jan"); s["B1"].SetNumber(100);
                            s["A2"].SetString("Feb"); s["B2"].SetNumber(150);
                        }
                        s.AddChart(type, "D1", "K15", "A1:A2", "B1:B2", title: "T-" + type);
                    }
                    wb.Save(path);
                }

                using var opened = Workbook.OpenOoxml(path);
                return opened.OpenXmlDocument!.WorkbookPart!.WorksheetParts
                    .SelectMany(wsPart => wsPart.GetPartsOfType<DrawingsPart>())
                    .SelectMany(dp => dp.WorksheetDrawing!.Elements<XDR.TwoCellAnchor>()
                        .Select(anchor => (Drawings: dp, Anchor: anchor)))
                    .Select(x =>
                    {
                        string rid = x.Anchor.GetFirstChild<XDR.GraphicFrame>()!
                            .Graphic!.GraphicData!.GetFirstChild<C.ChartReference>()!.Id!.Value!;
                        var space = ((ChartPart)x.Drawings.GetPartById(rid)).ChartSpace!;
                        var plotArea = space.GetFirstChild<C.Chart>()!.PlotArea!;
                        var plot = plotArea.ChildElements.Single(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
                        var from = x.Anchor.FromMarker!;
                        var to = x.Anchor.ToMarker!;
                        return new ChartObs(
                            plot.LocalName,
                            plot.Elements<C.BarDirection>().SingleOrDefault()?.Val?.InnerText,
                            (int.Parse(from.ColumnId!.Text, System.Globalization.CultureInfo.InvariantCulture),
                             int.Parse(from.RowId!.Text, System.Globalization.CultureInfo.InvariantCulture)),
                            (int.Parse(to.ColumnId!.Text, System.Globalization.CultureInfo.InvariantCulture),
                             int.Parse(to.RowId!.Text, System.Globalization.CultureInfo.InvariantCulture)),
                            space.Descendants<C.Title>().SingleOrDefault()?.Descendants<A.Text>().Single().Text,
                            plot.Descendants<C.Formula>().Select(f => (string?)f.Text).ToArray(),
                            new[]
                            {
                                plot.Descendants<C.StringCache>().SelectMany(c => c.Elements<C.StringPoint>())
                                    .Select(p => (p.Index!.Value, p.NumericValue!.Text!)).ToArray(),
                                plot.Descendants<C.NumberingCache>().SelectMany(c => c.Elements<C.NumericPoint>())
                                    .Select(p => (p.Index!.Value, p.NumericValue!.Text!)).ToArray(),
                            },
                            plot.Descendants<C.DataPoint>()
                                .Select(d => (d.Index!.Val!.Value,
                                    d.ChartShapeProperties!.Descendants<A.SchemeColor>().Single().Val!.InnerText!))
                                .OrderBy(t => t.Item1).ToArray());
                    })
                    .OrderBy(o => o.Plot, StringComparer.Ordinal)
                    .ThenBy(o => o.BarDir, StringComparer.Ordinal)
                    .ThenBy(o => o.Title, StringComparer.Ordinal)
                    .ToArray();
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }

        var npoi = EmitAndProject(() => Workbook.Create());
        var sdk = EmitAndProject(() => Workbook.CreateOoxml());
        sdk.Should().BeEquivalentTo(npoi, o => o.WithStrictOrdering());
    }
}
