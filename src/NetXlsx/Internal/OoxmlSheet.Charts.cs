// I-82 engine swap — Open XML SDK-backed charts (slice 8, the I-75 surface).
//
// A chart is a ChartPart hung off the sheet's DrawingsPart, referenced by r:id from
// an xdr:graphicFrame inside a twoCellAnchor in the shared xdr:wsDr root
// (GetOrCreateDrawing — the worksheet <drawing> child is already schema-ordered by
// the pictures sub-slice). The anchor end cell is EXCLUSIVE (NPOI:
// drawing.CreateAnchor(..., c2, r2) — the picture/shape convention, not the
// connector's inclusive one; quirk #10 checked, not assumed).
//
// The chart XML mirrors the NPOI engine's emission (the parity oracle, quirk #11),
// captured from xl/charts/chart1.xml for all six ChartTypes before implementing:
//
//   chartSpace: roundedCorners(0) > chart(title? > plotArea > plotVisOnly(1)) >
//   printSettings(headerFooter + pageMargins + pageSetup).
//   Line:    lineChart(grouping standard, varyColors 0, ser(idx, order,
//            marker(symbol none), cat strRef, val numRef), axId 1, axId 2)
//   Bar/Col: barChart(barDir bar|col, grouping clustered, varyColors 0,
//            ser(idx, order, invertIfNegative 0, cat, val), axId 1, axId 2)
//   Pie:     pieChart(varyColors 0, ser(idx, order, dPt x N with accent1..6
//            cycling solid fills, cat, val)) — no axId children
//   Scatter: scatterChart(scatterStyle lineMarker, varyColors 0, ser(idx, order,
//            xVal numRef, yVal numRef), axId 1, axId 2)
//   Area:    areaChart(varyColors 0, ser(idx, order, cat, val), axId 1, axId 2)
//   Axes:    valAx(axId 2, left, crossBetween midCat) then catAx(axId 1, bottom),
//            in that plotArea order.
//   Caches:  strCache/numCache snapshot the referenced cells at AddChart time.
//            ptCount = the full range size; only type-matching cells get a <c:pt>
//            (string cells in strCache; numeric/date cells in numCache — empty,
//            bool, formula, and mismatched cells are skipped), exactly as NPOI's
//            DataSources behave. A single-cell range collapses to "Sheet!$A$1".
//
// Documented conformance-positive divergences from NPOI (quirk #14 — the schema
// gate + Excel rendering win over NPOI byte-parity):
//   - NPOI writes the pie ser's dPt list AFTER cat/val; CT_PieSer requires dPt
//     BEFORE cat. The SDK engine emits schema order.
//   - NPOI emits a dangling catAx/valAx pair on pie charts (pieChart has no axId
//     children referencing them). The SDK engine omits axes on pie.
//   - NPOI pairs scatterChart's x axis with a catAx; an ECMA-376 scatter chart
//     plots two value axes. The SDK engine emits valAx (x, bottom) + valAx (y,
//     left) so the x axis scales numerically as a scatter should.
//   - NPOI's graphicFrame cNvPr is id=0 name="Diagramm0"; ST_DrawingElementId
//     must be unique-nonzero within the drawing (quirk #9), so the SDK engine
//     allocates NextShapeId like every other drawing kind.
//   - NPOI emits editAs="twoCell" explicitly; that is the attribute's schema
//     default, so the SDK engine omits it (same as the pictures sub-slice).

using System;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    private const string ChartGraphicDataUri =
        "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public IChart AddChart(ChartType type, string startCell, string endCell,
        string categoryRange, string valueRange, string? title = null)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(startCell);
        ArgumentNullException.ThrowIfNull(endCell);
        ArgumentNullException.ThrowIfNull(categoryRange);
        ArgumentNullException.ThrowIfNull(valueRange);

        var (r1, c1) = CellAddress.Parse(startCell);
        var (r2, c2) = CellAddress.Parse(endCell);
        var cat = CellAddress.ParseRange(categoryRange);
        var val = CellAddress.ParseRange(valueRange);

        var (drawingsPart, root) = GetOrCreateDrawing();
        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = BuildChartSpace(type, cat, val, title);

        uint id = NextShapeId(root);
        var frame = new XDR.GraphicFrame(
            new XDR.NonVisualGraphicFrameProperties(
                new XDR.NonVisualDrawingProperties { Id = id, Name = $"Chart {id}" },
                new XDR.NonVisualGraphicFrameDrawingProperties()),
            new XDR.Transform(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 0L, Cy = 0L }),
            new A.Graphic(
                new A.GraphicData(
                    new C.ChartReference { Id = drawingsPart.GetIdOfPart(chartPart) })
                { Uri = ChartGraphicDataUri }))
        { Macro = string.Empty };

        // Chart anchor end cell is exclusive (NPOI: CreateAnchor(..., c2, r2) —
        // same convention as pictures/shapes; connectors differ, quirk #10).
        root.AppendChild(new XDR.TwoCellAnchor(
            Marker(c1 - 1, 0, r1 - 1, 0),
            ToMarker(c2, 0, r2, 0),
            frame,
            new XDR.ClientData()));

        return new OoxmlChart(_workbook, this, type, chartPart);
    }

    // ---- chartSpace construction --------------------------------------------

    private C.ChartSpace BuildChartSpace(ChartType type,
        (int Row1, int Col1, int Row2, int Col2) cat,
        (int Row1, int Col1, int Row2, int Col2) val,
        string? title)
    {
        var plotArea = new C.PlotArea(new C.Layout());
        plotArea.Append(BuildPlotGroup(type, cat, val));
        if (type != ChartType.Pie)
        {
            // valAx then catAx, matching the NPOI engine's plotArea order.
            plotArea.Append(BuildValueAxis(
                axisId: 2, crossingAxisId: 1, C.AxisPositionValues.Left, crossBetween: true));
            plotArea.Append(type == ChartType.Scatter
                ? BuildValueAxis(
                    axisId: 1, crossingAxisId: 2, C.AxisPositionValues.Bottom, crossBetween: false)
                : BuildCategoryAxis(axisId: 1, crossingAxisId: 2));
        }

        var chart = new C.Chart();
        if (title is not null) chart.Append(OoxmlChart.BuildTitle(title));
        chart.Append(plotArea);
        chart.Append(new C.PlotVisibleOnly { Val = true });

        var chartSpace = new C.ChartSpace(
            new C.RoundedCorners { Val = false },
            chart,
            BuildPrintSettings());
        // Declare a/r at the root (the shape Excel and NPOI write) instead of
        // letting the SDK inline xmlns:a on each DrawingML descendant.
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        return chartSpace;
    }

    private OpenXmlCompositeElement BuildPlotGroup(ChartType type,
        (int Row1, int Col1, int Row2, int Col2) cat,
        (int Row1, int Col1, int Row2, int Col2) val)
    {
        switch (type)
        {
            case ChartType.Line:
                return new C.LineChart(
                    new C.Grouping { Val = C.GroupingValues.Standard },
                    new C.VaryColors { Val = false },
                    new C.LineChartSeries(
                        new C.Index { Val = 0U },
                        new C.Order { Val = 0U },
                        new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.None }),
                        new C.CategoryAxisData(BuildStringReference(cat)),
                        new C.Values(BuildNumberReference(val))),
                    new C.AxisId { Val = 1U },
                    new C.AxisId { Val = 2U });

            case ChartType.Bar:
            case ChartType.Column:
                return new C.BarChart(
                    new C.BarDirection
                    {
                        Val = type == ChartType.Bar
                            ? C.BarDirectionValues.Bar
                            : C.BarDirectionValues.Column,
                    },
                    new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                    new C.VaryColors { Val = false },
                    new C.BarChartSeries(
                        new C.Index { Val = 0U },
                        new C.Order { Val = 0U },
                        new C.InvertIfNegative { Val = false },
                        new C.CategoryAxisData(BuildStringReference(cat)),
                        new C.Values(BuildNumberReference(val))),
                    new C.AxisId { Val = 1U },
                    new C.AxisId { Val = 2U });

            case ChartType.Pie:
            {
                // I-89 lazy default-theme: the accent-cycled slice fills below
                // are theme-indexed (schemeClr accent1..6) — embed the default
                // theme so they resolve consumer-independently. The other chart
                // types emit no scheme colors and trigger nothing.
                _workbook.EnsureThemePart();

                var ser = new C.PieChartSeries(
                    new C.Index { Val = 0U },
                    new C.Order { Val = 0U });
                // One accent-cycled solid fill per slice, like NPOI — but emitted
                // in CT_PieSer schema order: dPt BEFORE cat/val (NPOI writes the
                // dPt list after val, which is schema-nonconformant; quirk #14).
                int points = RangeSize(cat);
                for (int i = 0; i < points; i++)
                {
                    ser.Append(new C.DataPoint(
                        new C.Index { Val = (uint)i },
                        new C.ChartShapeProperties(
                            new A.SolidFill(new A.SchemeColor { Val = AccentFor(i) }))));
                }
                ser.Append(new C.CategoryAxisData(BuildStringReference(cat)));
                ser.Append(new C.Values(BuildNumberReference(val)));
                return new C.PieChart(new C.VaryColors { Val = false }, ser);
            }

            case ChartType.Scatter:
                return new C.ScatterChart(
                    new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
                    new C.VaryColors { Val = false },
                    new C.ScatterChartSeries(
                        new C.Index { Val = 0U },
                        new C.Order { Val = 0U },
                        new C.XValues(BuildNumberReference(cat)),
                        new C.YValues(BuildNumberReference(val))),
                    new C.AxisId { Val = 1U },
                    new C.AxisId { Val = 2U });

            default: // ChartType.Area
                return new C.AreaChart(
                    new C.VaryColors { Val = false },
                    new C.AreaChartSeries(
                        new C.Index { Val = 0U },
                        new C.Order { Val = 0U },
                        new C.CategoryAxisData(BuildStringReference(cat)),
                        new C.Values(BuildNumberReference(val))),
                    new C.AxisId { Val = 1U },
                    new C.AxisId { Val = 2U });
        }
    }

    // ---- data references (formula + value cache) ----------------------------

    private C.StringReference BuildStringReference((int Row1, int Col1, int Row2, int Col2) range)
    {
        var cache = new C.StringCache(new C.PointCount { Val = (uint)RangeSize(range) });
        uint idx = 0;
        for (int row = range.Row1; row <= range.Row2; row++)
        {
            for (int col = range.Col1; col <= range.Col2; col++, idx++)
            {
                var cell = this[row, col];
                if (cell.Kind == CellKind.String)
                    cache.Append(new C.StringPoint(new C.NumericValue(cell.GetString())) { Index = idx });
            }
        }
        return new C.StringReference(new C.Formula(FormulaRef(range)), cache);
    }

    private C.NumberReference BuildNumberReference((int Row1, int Col1, int Row2, int Col2) range)
    {
        var cache = new C.NumberingCache(new C.PointCount { Val = (uint)RangeSize(range) });
        uint idx = 0;
        for (int row = range.Row1; row <= range.Row2; row++)
        {
            for (int col = range.Col1; col <= range.Col2; col++, idx++)
            {
                var cell = this[row, col];
                // A date cell is a numeric serial; bool/string/formula cells are
                // skipped (NPOI's numeric range source includes CellType.Numeric only).
                if (cell.Kind is CellKind.Number or CellKind.Date)
                {
                    cache.Append(new C.NumericPoint(
                        new C.NumericValue(cell.GetNumber()!.Value.ToString(CultureInfo.InvariantCulture)))
                    { Index = idx });
                }
            }
        }
        return new C.NumberReference(new C.Formula(FormulaRef(range)), cache);
    }

    private static int RangeSize((int Row1, int Col1, int Row2, int Col2) range) =>
        (range.Row2 - range.Row1 + 1) * (range.Col2 - range.Col1 + 1);

    // "Data!$A$1:$A$4" — sheet-qualified absolute reference; a single-cell range
    // collapses to "Data!$A$1" (both forms matched against the NPOI oracle).
    private string FormulaRef((int Row1, int Col1, int Row2, int Col2) range)
    {
        string start = AbsoluteRef(range.Row1, range.Col1);
        string body = range.Row1 == range.Row2 && range.Col1 == range.Col2
            ? start
            : $"{start}:{AbsoluteRef(range.Row2, range.Col2)}";
        return $"{QuoteSheetName(Name)}!{body}";
    }

    // ---- axes / fixed trailing blocks ---------------------------------------

    private static C.ValueAxis BuildValueAxis(uint axisId, uint crossingAxisId,
        C.AxisPositionValues position, bool crossBetween)
    {
        var axis = new C.ValueAxis(
            new C.AxisId { Val = axisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = position },
            new C.MajorTickMark { Val = C.TickMarkValues.Cross },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = crossingAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero });
        // crossBetween is meaningful on the y axis (NPOI emits midCat there); a
        // scatter x axis omits it.
        if (crossBetween)
            axis.Append(new C.CrossBetween { Val = C.CrossBetweenValues.MidpointCategory });
        return axis;
    }

    private static C.CategoryAxis BuildCategoryAxis(uint axisId, uint crossingAxisId) => new(
        new C.AxisId { Val = axisId },
        new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
        new C.Delete { Val = false },
        new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
        new C.MajorTickMark { Val = C.TickMarkValues.Cross },
        new C.MinorTickMark { Val = C.TickMarkValues.None },
        new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
        new C.CrossingAxis { Val = crossingAxisId },
        new C.Crosses { Val = C.CrossesValues.AutoZero });

    // NPOI's fixed printSettings block, reproduced verbatim.
    private static C.PrintSettings BuildPrintSettings() => new(
        new C.HeaderFooter(),
        new C.PageMargins
        {
            Left = 0.7,
            Right = 0.7,
            Top = 0.75,
            Bottom = 0.75,
            Header = 0.3,
            Footer = 0.3,
        },
        new C.PageSetup
        {
            PaperSize = 1U,
            FirstPageNumber = 1,
            Orientation = C.PageSetupOrientationValues.Default,
            BlackAndWhite = false,
            Draft = false,
            UseFirstPageNumber = false,
            HorizontalDpi = 600,
            VerticalDpi = 600,
            Copies = 1U,
        });

    private static A.SchemeColorValues AccentFor(int index) => (index % 6) switch
    {
        0 => A.SchemeColorValues.Accent1,
        1 => A.SchemeColorValues.Accent2,
        2 => A.SchemeColorValues.Accent3,
        3 => A.SchemeColorValues.Accent4,
        4 => A.SchemeColorValues.Accent5,
        _ => A.SchemeColorValues.Accent6,
    };
}
