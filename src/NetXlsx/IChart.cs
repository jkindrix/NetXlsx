namespace NetXlsx;

/// <summary>
/// Chart types supported by <see cref="ISheet.AddChart"/> (decision I-75).
/// </summary>
public enum ChartType
{
    /// <summary>Line chart.</summary>
    Line,
    /// <summary>Vertical bar chart.</summary>
    Bar,
    /// <summary>Column chart (horizontal bars).</summary>
    Column,
    /// <summary>Pie chart.</summary>
    Pie,
    /// <summary>Scatter (XY) chart.</summary>
    Scatter,
    /// <summary>Area chart.</summary>
    Area,
}

/// <summary>
/// A chart embedded on a sheet (decision I-75). Constructed via
/// <see cref="ISheet.AddChart"/>. For advanced customization (series
/// formatting, axis options, legend), reach through <see cref="Underlying"/>.
/// </summary>
public interface IChart
{
    /// <summary>The owning sheet.</summary>
    ISheet Sheet { get; }

    /// <summary>The chart type.</summary>
    ChartType Type { get; }

    /// <summary>Sets the chart title.</summary>
    void SetTitle(string title);

    /// <summary>
    /// Escape hatch — direct access to the underlying NPOI
    /// <c>XSSFChart</c>.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFChart Underlying { get; }
}
