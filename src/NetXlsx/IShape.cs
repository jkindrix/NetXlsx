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
    /// Escape hatch — direct access to the underlying NPOI
    /// <c>XSSFSimpleShape</c>.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFSimpleShape Underlying { get; }
}

/// <summary>
/// Shape types supported by <see cref="ISheet.AddShape"/> (decision I-74).
/// Maps to NPOI's <c>ShapeTypes</c> enum. Only the most-common shapes
/// are surfaced; callers needing exotic shapes (arrows, callouts, stars)
/// reach through <see cref="ISheet.Underlying"/>.
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

    /// <summary>
    /// Escape hatch — direct access to the underlying NPOI
    /// <c>XSSFConnector</c>.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFConnector Underlying { get; }
}
