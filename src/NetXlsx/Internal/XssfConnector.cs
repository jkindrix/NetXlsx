using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfConnector : IConnector
{
    private readonly XssfSheet _sheet;
    private readonly XSSFConnector _underlying;
    private readonly ConnectorType _type;

    internal XssfConnector(XssfSheet sheet, XSSFConnector underlying, ConnectorType type)
    {
        _sheet = sheet;
        _underlying = underlying;
        _type = type;
    }

    /// <summary>
    /// Wraps an existing NPOI connector (read path, decision I-81),
    /// deriving the <see cref="ConnectorType"/> from the preset geometry.
    /// Unknown geometries fall back to <see cref="ConnectorType.Straight"/>.
    /// </summary>
    internal static XssfConnector FromExisting(XssfSheet sheet, XSSFConnector conn)
    {
        var prst = conn.GetCTConnector()?.spPr?.prstGeom?.prst;
        var type = prst switch
        {
            ST_ShapeType.bentConnector2 or ST_ShapeType.bentConnector3
                or ST_ShapeType.bentConnector4 or ST_ShapeType.bentConnector5 => ConnectorType.Bent,
            ST_ShapeType.curvedConnector2 or ST_ShapeType.curvedConnector3
                or ST_ShapeType.curvedConnector4 or ST_ShapeType.curvedConnector5 => ConnectorType.Curved,
            _ => ConnectorType.Straight,
        };
        return new XssfConnector(sheet, conn, type);
    }

    public ISheet Sheet => _sheet;
    public ConnectorType Type => _type;
    public XSSFConnector Underlying => _underlying;

    // ---- Anchor (decision I-81) ---------------------------------------

    private XSSFClientAnchor Anchor => (XSSFClientAnchor)_underlying.GetAnchor();

    public string FromCell { get { var a = Anchor; return CellAddress.Format(a.Row1 + 1, a.Col1 + 1); } }
    public string ToCell   { get { var a = Anchor; return CellAddress.Format(a.Row2 + 1, a.Col2 + 1); } }
    public int Dx1 => Anchor.Dx1;
    public int Dy1 => Anchor.Dy1;
    public int Dx2 => Anchor.Dx2;
    public int Dy2 => Anchor.Dy2;

    // ---- Transform (flip) and arrowheads (decision I-81) --------------

    // CT_ShapeProperties used on the connector's spPr lives in the
    // Spreadsheet namespace (CT_Connector.spPr is the spreadsheet variant).
    private NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties? SpPr =>
        _underlying.GetCTConnector()?.spPr;

    public bool FlipH => SpPr?.xfrm?.flipH ?? false;
    public bool FlipV => SpPr?.xfrm?.flipV ?? false;

    public ConnectorEnd HeadEnd => ToConnectorEnd(SpPr?.ln?.headEnd?.type);
    public ConnectorEnd TailEnd => ToConnectorEnd(SpPr?.ln?.tailEnd?.type);

    private static ConnectorEnd ToConnectorEnd(ST_LineEndType? t) => t switch
    {
        ST_LineEndType.triangle => ConnectorEnd.Triangle,
        ST_LineEndType.stealth => ConnectorEnd.Stealth,
        ST_LineEndType.diamond => ConnectorEnd.Diamond,
        ST_LineEndType.oval => ConnectorEnd.Oval,
        ST_LineEndType.arrow => ConnectorEnd.Arrow,
        _ => ConnectorEnd.None,
    };

    // ---- Explicit line color / width + theme-style refs (decision I-81)

    public Color? LineColor
    {
        get
        {
            var rgb = SpPr?.ln?.solidFill?.srgbClr?.val;
            if (rgb is { Length: >= 3 })
                return Color.FromRgb(rgb[0], rgb[1], rgb[2]);
            return null;
        }
    }

    public string? LineSchemeColor
    {
        get
        {
            // Prefer the explicit ln.solidFill.schemeClr (rare); fall back
            // to the connector style's lnRef.schemeClr (the common case).
            var explicitName = SpPr?.ln?.solidFill?.schemeClr?.val.ToString();
            if (!string.IsNullOrEmpty(explicitName)) return explicitName;
            return _underlying.GetCTConnector()?.style?.lnRef?.schemeClr?.val.ToString();
        }
    }

    public double? LineWidthPoints
    {
        get
        {
            // NPOI's XSSFConnector.LineWidth setter populates ln.w but
            // leaves wSpecified=false, so gate on the value, not the flag.
            // 0 is OOXML's "unset" default for the integer EMU width.
            var ln = SpPr?.ln;
            if (ln is null || ln.w <= 0) return null;
            return ln.w / 12700.0;
        }
    }

    public int? LineStyleRefIndex
    {
        get
        {
            var lnRef = _underlying.GetCTConnector()?.style?.lnRef;
            return lnRef is null ? null : (int)lnRef.idx;
        }
    }
}
