// I-82 engine swap — Open XML SDK-backed IChart (charts slice, the I-75 surface).
//
// A chart on the SDK engine. IChart's surface is intentionally minimal (Sheet, Type,
// SetTitle, Underlying) — series/axis customization reaches through the escape hatch,
// which on this engine is IWorkbook.OpenXmlDocument; the NPOI escape hatch
// (Underlying -> XSSFChart) throws NotSupportedException, the same divergence as
// OoxmlPicture/OoxmlShape/OoxmlConnector/OoxmlTable: there is no NPOI object behind
// the SDK engine. The chart XML itself is built by OoxmlSheet.Charts.cs; this wrapper
// keeps the ChartPart so SetTitle can rewrite the c:title in place.

using System;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace NetXlsx;

internal sealed class OoxmlChart : IChart
{
    private readonly OoxmlWorkbook _workbook;
    private readonly OoxmlSheet _sheet;
    private readonly ChartType _type;
    private readonly ChartPart _part;

    internal OoxmlChart(OoxmlWorkbook workbook, OoxmlSheet sheet, ChartType type, ChartPart part)
    {
        _workbook = workbook;
        _sheet = sheet;
        _type = type;
        _part = part;
    }

    public ISheet Sheet { get { _workbook.ThrowIfDisposed(); return _sheet; } }
    public ChartType Type { get { _workbook.ThrowIfDisposed(); return _type; } }

    public void SetTitle(string title)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(title);
        var chart = _part.ChartSpace!.GetFirstChild<C.Chart>()!;
        chart.RemoveAllChildren<C.Title>();
        // CT_Chart is a strict sequence and <c:title> is its first member.
        chart.InsertAt(BuildTitle(title), 0);
    }

    // The NPOI engine's title shape (the parity oracle): c:tx/c:rich with a
    // zero-inset bodyPr and a single unformatted run. NPOI's SetTitle replaces
    // the previous title outright, so a re-set leaves exactly one c:title.
    internal static C.Title BuildTitle(string text) => new(
        new C.ChartText(new C.RichText(
            new A.BodyProperties
            {
                LeftInset = 0,
                TopInset = 0,
                RightInset = 0,
                BottomInset = 0,
                RightToLeftColumns = false,
            },
            new A.Paragraph(new A.Run(new A.Text(text))))));

    // Escape-hatch divergence (I-82): no NPOI chart exists on the SDK engine.
    public NPOI.XSSF.UserModel.XSSFChart Underlying => throw new NotSupportedException(
        "IChart.Underlying (NPOI XSSFChart) is not available on the Open XML SDK " +
        "engine (I-82). Use IWorkbook.OpenXmlDocument for the SDK escape hatch.");
}
