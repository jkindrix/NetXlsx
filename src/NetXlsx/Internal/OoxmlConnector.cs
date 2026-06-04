// I-82 engine swap — Open XML SDK-backed IConnector (drawings slice, shapes/connectors).
//
// An immutable snapshot of one anchored connector (xdr:cxnSp) on the SDK engine.
// FromElement parses the twoCellAnchor + cxnSp once; both OoxmlSheet.AddConnector and
// OoxmlSheet.Connectors flow through it, so the just-created and the re-read views can
// never drift. The read surface mirrors the NPOI engine's XssfConnector contract
// exactly (decision I-81), the parity oracle for this geometry surface:
//
//   Type            derived from spPr/prstGeom/@prst (straightConnector1 -> Straight,
//                   bentConnector* -> Bent, curvedConnector* -> Curved, else Straight),
//                   matching XssfConnector.FromExisting.
//   FromCell/ToCell A1 corners; the connector anchor's end cell is INCLUSIVE (the add
//                   path stores c2-1/r2-1), so a same-cell connector reports
//                   ToCell == FromCell. Dx1/Dy1/Dx2/Dy2 are the per-end EMU offsets.
//   FlipH/FlipV     spPr/xfrm/@flipH /@flipV.
//   Head/TailEnd    spPr/ln/headEnd /tailEnd /@type.
//   LineColor       spPr/ln/solidFill/srgbClr (explicit RGB), else null.
//   LineSchemeColor explicit ln solidFill schemeClr, else style/lnRef/schemeClr
//                   (="accent1" for an engine-created connector).
//   LineWidthPoints spPr/ln/@w / 12700, null when unset (0 is OOXML's unset default).
//   LineStyleRefIndex style/lnRef/@idx (=1 for an engine-created connector).
//
// The NPOI escape hatch (Underlying -> XSSFConnector) throws NotSupportedException,
// the same divergence as OoxmlPicture/OoxmlShape.

using System;
using DocumentFormat.OpenXml;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;

namespace NetXlsx;

internal sealed class OoxmlConnector : IConnector
{
    private const double EmuPerPoint = 12700.0;

    private readonly OoxmlWorkbook _workbook;
    private readonly OoxmlSheet _sheet;
    private readonly ConnectorType _type;
    private readonly string _fromCell, _toCell;
    private readonly int _dx1, _dy1, _dx2, _dy2;
    private readonly bool _flipH, _flipV;
    private readonly ConnectorEnd _headEnd, _tailEnd;
    private readonly Color? _lineColor;
    private readonly string? _lineSchemeColor;
    private readonly double? _lineWidthPoints;
    private readonly int? _lineStyleRefIndex;
    private readonly XDR.ConnectionShape _cxn;

    private OoxmlConnector(
        OoxmlWorkbook workbook, OoxmlSheet sheet, ConnectorType type,
        string fromCell, string toCell, int dx1, int dy1, int dx2, int dy2,
        bool flipH, bool flipV, ConnectorEnd headEnd, ConnectorEnd tailEnd,
        Color? lineColor, string? lineSchemeColor, double? lineWidthPoints, int? lineStyleRefIndex,
        XDR.ConnectionShape cxn)
    {
        _workbook = workbook;
        _sheet = sheet;
        _type = type;
        _fromCell = fromCell; _toCell = toCell;
        _dx1 = dx1; _dy1 = dy1; _dx2 = dx2; _dy2 = dy2;
        _flipH = flipH; _flipV = flipV;
        _headEnd = headEnd; _tailEnd = tailEnd;
        _lineColor = lineColor;
        _lineSchemeColor = lineSchemeColor;
        _lineWidthPoints = lineWidthPoints;
        _lineStyleRefIndex = lineStyleRefIndex;
        _cxn = cxn;
    }

    internal static OoxmlConnector FromElement(
        OoxmlWorkbook workbook, OoxmlSheet sheet, XDR.TwoCellAnchor anchor, XDR.ConnectionShape cxn)
    {
        var from = anchor.GetFirstChild<XDR.FromMarker>()!;
        var to = anchor.GetFirstChild<XDR.ToMarker>()!;
        int fc = OoxmlSheet.ParseMarker(from.ColumnId), fco = OoxmlSheet.ParseMarker(from.ColumnOffset);
        int fr = OoxmlSheet.ParseMarker(from.RowId), fro = OoxmlSheet.ParseMarker(from.RowOffset);
        int tc = OoxmlSheet.ParseMarker(to.ColumnId), tco = OoxmlSheet.ParseMarker(to.ColumnOffset);
        int tr = OoxmlSheet.ParseMarker(to.RowId), tro = OoxmlSheet.ParseMarker(to.RowOffset);

        var spPr = cxn.ShapeProperties;
        var preset = spPr?.GetFirstChild<A.PresetGeometry>()?.Preset;
        var type = ToConnectorType(preset);

        var xfrm = spPr?.Transform2D;
        bool flipH = xfrm?.HorizontalFlip?.Value ?? false;
        bool flipV = xfrm?.VerticalFlip?.Value ?? false;

        var ln = spPr?.GetFirstChild<A.Outline>();
        var headEnd = ToConnectorEnd(ln?.GetFirstChild<A.HeadEnd>()?.Type);
        var tailEnd = ToConnectorEnd(ln?.GetFirstChild<A.TailEnd>()?.Type);

        var solidFill = ln?.GetFirstChild<A.SolidFill>();
        Color? lineColor = null;
        if (solidFill?.RgbColorModelHex?.Val?.Value is string hex && TryParseHex(hex, out var rgb))
            lineColor = rgb;

        // Explicit ln scheme color first, then the style block's lnRef (the common case).
        string? lineSchemeColor = solidFill?.SchemeColor?.Val?.InnerText;
        if (string.IsNullOrEmpty(lineSchemeColor))
            lineSchemeColor = cxn.ShapeStyle?.LineReference?.SchemeColor?.Val?.InnerText;

        // <a:ln @w> is EMU; 0/unset means "no explicit width" (theme-driven).
        double? lineWidthPoints = null;
        if (ln?.Width?.Value is int w && w > 0)
            lineWidthPoints = w / EmuPerPoint;

        int? lineStyleRefIndex = null;
        if (cxn.ShapeStyle?.LineReference?.Index?.Value is uint idx)
            lineStyleRefIndex = (int)idx;

        return new OoxmlConnector(workbook, sheet, type,
            CellAddress.Format(fr + 1, fc + 1), CellAddress.Format(tr + 1, tc + 1),
            fco, fro, tco, tro, flipH, flipV, headEnd, tailEnd,
            lineColor, lineSchemeColor, lineWidthPoints, lineStyleRefIndex, cxn);
    }

    public ISheet Sheet { get { _workbook.ThrowIfDisposed(); return _sheet; } }
    public ConnectorType Type { get { _workbook.ThrowIfDisposed(); return _type; } }
    public string FromCell { get { _workbook.ThrowIfDisposed(); return _fromCell; } }
    public string ToCell { get { _workbook.ThrowIfDisposed(); return _toCell; } }
    public int Dx1 { get { _workbook.ThrowIfDisposed(); return _dx1; } }
    public int Dy1 { get { _workbook.ThrowIfDisposed(); return _dy1; } }
    public int Dx2 { get { _workbook.ThrowIfDisposed(); return _dx2; } }
    public int Dy2 { get { _workbook.ThrowIfDisposed(); return _dy2; } }
    public bool FlipH { get { _workbook.ThrowIfDisposed(); return _flipH; } }
    public bool FlipV { get { _workbook.ThrowIfDisposed(); return _flipV; } }
    public ConnectorEnd HeadEnd { get { _workbook.ThrowIfDisposed(); return _headEnd; } }
    public ConnectorEnd TailEnd { get { _workbook.ThrowIfDisposed(); return _tailEnd; } }
    public Color? LineColor { get { _workbook.ThrowIfDisposed(); return _lineColor; } }
    public string? LineSchemeColor { get { _workbook.ThrowIfDisposed(); return _lineSchemeColor; } }
    public double? LineWidthPoints { get { _workbook.ThrowIfDisposed(); return _lineWidthPoints; } }
    public int? LineStyleRefIndex { get { _workbook.ThrowIfDisposed(); return _lineStyleRefIndex; } }

    // Escape hatch (#32 / I-82): the live xdr:cxnSp element. Disposal first.
    public XDR.ConnectionShape Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _cxn; }
    }

    // ---- Mappings (inverse of OoxmlSheet.ToPreset / ToLineEnd) ---------------

    private static ConnectorType ToConnectorType(EnumValue<A.ShapeTypeValues>? preset)
    {
        if (preset is null) return ConnectorType.Straight;
        if (preset == A.ShapeTypeValues.BentConnector2 || preset == A.ShapeTypeValues.BentConnector3
            || preset == A.ShapeTypeValues.BentConnector4 || preset == A.ShapeTypeValues.BentConnector5)
            return ConnectorType.Bent;
        if (preset == A.ShapeTypeValues.CurvedConnector2 || preset == A.ShapeTypeValues.CurvedConnector3
            || preset == A.ShapeTypeValues.CurvedConnector4 || preset == A.ShapeTypeValues.CurvedConnector5)
            return ConnectorType.Curved;
        return ConnectorType.Straight;
    }

    private static ConnectorEnd ToConnectorEnd(EnumValue<A.LineEndValues>? type)
    {
        if (type is null) return ConnectorEnd.None;
        if (type == A.LineEndValues.Triangle) return ConnectorEnd.Triangle;
        if (type == A.LineEndValues.Stealth) return ConnectorEnd.Stealth;
        if (type == A.LineEndValues.Diamond) return ConnectorEnd.Diamond;
        if (type == A.LineEndValues.Oval) return ConnectorEnd.Oval;
        if (type == A.LineEndValues.Arrow) return ConnectorEnd.Arrow;
        return ConnectorEnd.None;
    }

    private static bool TryParseHex(string hex, out Color color)
    {
        color = default;
        if (hex.Length == 8) hex = hex.Substring(2); // strip leading AA in AARRGGBB
        if (hex.Length != 6) return false;
        if (byte.TryParse(hex.AsSpan(0, 2), System.Globalization.NumberStyles.HexNumber, null, out var r)
            && byte.TryParse(hex.AsSpan(2, 2), System.Globalization.NumberStyles.HexNumber, null, out var g)
            && byte.TryParse(hex.AsSpan(4, 2), System.Globalization.NumberStyles.HexNumber, null, out var b))
        {
            color = Color.FromRgb(r, g, b);
            return true;
        }
        return false;
    }
}
