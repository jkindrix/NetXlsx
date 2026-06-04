namespace NetXlsx;

/// <summary>
/// A shape drawn on a sheet (decision I-74). Constructed via
/// <see cref="ISheet.AddShape"/>. For advanced manipulation (rotation,
/// line width, gradients), reach through <see cref="Underlying"/>.
/// </summary>
public interface IShape
{
    /// <summary>The owning sheet.</summary>
    ISheet Sheet { get; }

    /// <summary>The shape type.</summary>
    ShapeType Type { get; }

    /// <summary>
    /// Escape hatch — direct access to the underlying Open XML SDK
    /// <see cref="DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape"/>
    /// element (I-82). Same contract as <see cref="IWorkbook.Underlying"/>.
    /// </summary>
    DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape Underlying { get; }
}

/// <summary>
/// Shape types supported by <see cref="ISheet.AddShape"/> (decision I-74).
/// Values are the OOXML <c>ST_ShapeType</c> preset-geometry ordinals.
/// Only the most-common shapes are surfaced; callers needing exotic
/// shapes (arrows, callouts, stars) reach through
/// <see cref="ISheet.Underlying"/>.
/// </summary>
public enum ShapeType
{
    /// <summary>A rectangle.</summary>
    Rectangle = 5,
    /// <summary>A rounded rectangle.</summary>
    RoundedRectangle = 26,
    /// <summary>An ellipse / circle.</summary>
    Ellipse = 35,
    /// <summary>A straight line.</summary>
    Line = 1,
    /// <summary>A triangle.</summary>
    Triangle = 3,
    /// <summary>A diamond.</summary>
    Diamond = 6,
}

/// <summary>
/// Connector types for <see cref="ISheet.AddConnector"/> (decisions I-79, I-80).
/// Values are <c>ST_ShapeType</c> ordinals for the corresponding preset
/// connector geometry.
/// </summary>
public enum ConnectorType
{
    /// <summary>A straight line connector (<c>straightConnector1</c>).</summary>
    Straight = 96,
    /// <summary>A bent (right-angle) connector (<c>bentConnector3</c>).</summary>
    Bent = 98,
    /// <summary>A curved connector (<c>curvedConnector3</c>).</summary>
    Curved = 102,
}

/// <summary>
/// Line-end (arrowhead) decorations for connector ends (decision I-80).
/// Maps to OOXML <c>ST_LineEndType</c>.
/// </summary>
public enum ConnectorEnd
{
    /// <summary>No decoration (plain line end).</summary>
    None = 0,
    /// <summary>A filled triangular arrowhead.</summary>
    Triangle,
    /// <summary>A "stealth" (concave) arrowhead.</summary>
    Stealth,
    /// <summary>A diamond.</summary>
    Diamond,
    /// <summary>An oval.</summary>
    Oval,
    /// <summary>An open arrow (the Excel "arrow" line end).</summary>
    Arrow,
}

/// <summary>
/// A connector (line or arrow) drawn on a sheet (decisions I-79, I-80).
/// Constructed via <see cref="ISheet.AddConnector"/>. For advanced
/// manipulation, reach through <see cref="Underlying"/>.
/// </summary>
public interface IConnector
{
    /// <summary>The owning sheet.</summary>
    ISheet Sheet { get; }

    /// <summary>The connector type.</summary>
    ConnectorType Type { get; }

    /// <summary>Start anchor cell in A1 notation (decision I-81).</summary>
    string FromCell { get; }
    /// <summary>End anchor cell in A1 notation (decision I-81).</summary>
    string ToCell { get; }

    /// <summary>EMU x-offset of the start point within <see cref="FromCell"/> (I-81).</summary>
    int Dx1 { get; }
    /// <summary>EMU y-offset of the start point within <see cref="FromCell"/> (I-81).</summary>
    int Dy1 { get; }
    /// <summary>EMU x-offset of the end point within <see cref="ToCell"/> (I-81).</summary>
    int Dx2 { get; }
    /// <summary>EMU y-offset of the end point within <see cref="ToCell"/> (I-81).</summary>
    int Dy2 { get; }

    /// <summary>Whether the connector is flipped horizontally (decision I-81).</summary>
    bool FlipH { get; }
    /// <summary>Whether the connector is flipped vertically (decision I-81).</summary>
    bool FlipV { get; }

    /// <summary>Arrowhead decoration on the start (head) end (decision I-81).</summary>
    ConnectorEnd HeadEnd { get; }
    /// <summary>Arrowhead decoration on the end (tail) end (decision I-81).</summary>
    ConnectorEnd TailEnd { get; }

    /// <summary>
    /// Explicit line color (<c>spPr/ln/solidFill/srgbClr</c>) if set on this
    /// connector, otherwise <c>null</c> (decision I-81). A null result with
    /// a non-null <see cref="LineSchemeColor"/> means the color comes from
    /// the theme — resolve via <see cref="IWorkbook.ResolveThemeColor(string, double)"/>.
    /// </summary>
    Color? LineColor { get; }

    /// <summary>
    /// Scheme-color name (e.g. <c>"dk1"</c>, <c>"accent1"</c>) referenced by
    /// either the explicit line fill or the connector's style
    /// <c>lnRef</c>, if any (decision I-81). Resolve via
    /// <see cref="IWorkbook.ResolveThemeColor(string, double)"/>.
    /// </summary>
    string? LineSchemeColor { get; }

    /// <summary>
    /// Explicit line width in points (<c>spPr/ln/@w</c> / 12700) if set,
    /// otherwise <c>null</c> (decision I-81). A null result with a non-null
    /// <see cref="LineStyleRefIndex"/> means width comes from the theme —
    /// resolve via <see cref="IWorkbook.GetThemeLineWidthEmu(int)"/>.
    /// </summary>
    double? LineWidthPoints { get; }

    /// <summary>
    /// 1-based index into the theme's <c>fmtScheme/lnStyleLst</c> from the
    /// connector's <c>style/lnRef/@idx</c>, if any (decision I-81).
    /// </summary>
    int? LineStyleRefIndex { get; }

    /// <summary>
    /// Escape hatch — direct access to the underlying Open XML SDK
    /// <see cref="DocumentFormat.OpenXml.Drawing.Spreadsheet.ConnectionShape"/>
    /// element (I-82). Same contract as <see cref="IWorkbook.Underlying"/>.
    /// </summary>
    DocumentFormat.OpenXml.Drawing.Spreadsheet.ConnectionShape Underlying { get; }
}
