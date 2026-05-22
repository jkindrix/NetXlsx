// TotalsRowFunction — per-column aggregation function for table
// totals rows (decision I-64). Mirrors the OOXML
// ST_TotalsRowFunction enum, exposed as a NetXlsx-namespaced public
// enum so callers don't take an NPOI dependency in their code.

namespace NetXlsx;

/// <summary>
/// Aggregation function applied to a column's totals-row cell
/// in an Excel Table (<see cref="ITable"/>). When set via
/// <see cref="ITable.SetColumnTotal(string, TotalsRowFunction)"/>,
/// NetXlsx writes the matching <c>SUBTOTAL</c> formula into the
/// underlying cell so the totals render in any conforming viewer.
/// <para>
/// SUBTOTAL is invoked in its 100-series form (function code
/// <c>10x</c>) to skip rows hidden by AutoFilter — matching Excel's
/// default behavior for table totals.
/// </para>
/// </summary>
public enum TotalsRowFunction
{
    /// <summary>No total — the cell is blank.</summary>
    None,
    /// <summary>SUM — <c>SUBTOTAL(109, ...)</c>.</summary>
    Sum,
    /// <summary>MIN — <c>SUBTOTAL(105, ...)</c>.</summary>
    Min,
    /// <summary>MAX — <c>SUBTOTAL(104, ...)</c>.</summary>
    Max,
    /// <summary>AVERAGE — <c>SUBTOTAL(101, ...)</c>.</summary>
    Average,
    /// <summary>COUNTA — <c>SUBTOTAL(103, ...)</c>.</summary>
    Count,
    /// <summary>COUNT — <c>SUBTOTAL(102, ...)</c>.</summary>
    CountNumbers,
    /// <summary>STDEV — <c>SUBTOTAL(107, ...)</c>.</summary>
    StdDev,
    /// <summary>VAR — <c>SUBTOTAL(110, ...)</c>.</summary>
    Var,
    /// <summary>Custom formula supplied via the second overload.</summary>
    Custom,
}
