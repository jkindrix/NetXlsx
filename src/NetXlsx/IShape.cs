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
/// Connector types for <see cref="ISheet.AddConnector"/> (decision I-79).
/// Maps to NPOI's connector shape types.
/// </summary>
public enum ConnectorType
{
    /// <summary>A straight line connector.</summary>
    Straight = 20,  // ShapeTypes.Line
    /// <summary>A bent (right-angle) connector.</summary>
    Bent = 32,
    /// <summary>A curved connector.</summary>
    Curved = 38,
}
