// Excel Tables (ListObject) per design §6.4.1 / decision I-51.
// Tables are sheet-scoped structured ranges with a header row, optional
// totals row, optional style, and (always-on in OOXML) AutoFilter.

namespace NetXlsx;

/// <summary>
/// A structured Excel table (OOXML <c>ListObject</c>) over a range on
/// a sheet. Tables differ from named ranges in that they are dynamic —
/// adding a row inside the table area grows the table — and they carry
/// a header row, optional totals row, and a style.
/// <para>
/// Construct via <see cref="ISheet.AddTable"/>; enumerate via
/// <see cref="ISheet.Tables"/>; remove via <see cref="ISheet.RemoveTable"/>.
/// </para>
/// </summary>
public interface ITable
{
    /// <summary>
    /// The table's codename (Excel-name rules: letters / digits /
    /// underscores; must start with a letter or underscore; must be
    /// unique workbook-wide, case-insensitive). Set when the table
    /// is added; not mutable on this surface in v1.1.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// The human-readable name shown in Excel's UI. Defaults to
    /// <see cref="Name"/>. Settable.
    /// </summary>
    string DisplayName { get; set; }

    /// <summary>
    /// The table's current A1 range, e.g. <c>"A1:D10"</c>. Read-only
    /// here — mutate by editing cells in the underlying sheet; Excel
    /// recomputes the range on next open. For programmatic resize,
    /// reach through <see cref="Underlying"/>.
    /// </summary>
    string Address { get; }

    /// <summary>The owning sheet.</summary>
    ISheet Sheet { get; }

    /// <summary>
    /// Column names, in order. Auto-derived from the header row at
    /// <see cref="ISheet.AddTable"/> time, and refreshed when read.
    /// Empty when the table has no columns.
    /// </summary>
    System.Collections.Generic.IReadOnlyList<string> ColumnNames { get; }

    /// <summary>
    /// Whether the table carries a totals row. Toggled via
    /// <see cref="AddTotalsRow"/> / <see cref="RemoveTotalsRow"/>
    /// (decision I-64).
    /// </summary>
    bool HasTotalsRow { get; }

    /// <summary>
    /// Adds a totals row beneath the table's data range (decision I-64).
    /// Extends the table's <see cref="Address"/> by one row down and
    /// sets <see cref="HasTotalsRow"/> to <c>true</c>. The totals row
    /// is initially blank; per-column aggregations are configured via
    /// <see cref="SetColumnTotal(string, TotalsRowFunction)"/>. No-op
    /// if a totals row is already present.
    /// </summary>
    void AddTotalsRow();

    /// <summary>
    /// Removes the totals row (decision I-64). Clears all per-column
    /// totals-row functions, shrinks the table range by one row, and
    /// sets <see cref="HasTotalsRow"/> to <c>false</c>. No-op if no
    /// totals row is present.
    /// </summary>
    void RemoveTotalsRow();

    /// <summary>
    /// Configures the totals-row cell for <paramref name="columnName"/>
    /// to compute <paramref name="function"/>. Writes the matching
    /// <c>SUBTOTAL</c> formula into the cell so the totals render in
    /// any conforming viewer (decision I-64).
    /// <para>
    /// For <see cref="TotalsRowFunction.Custom"/>, use the
    /// <see cref="SetColumnTotal(string, string)"/> overload.
    /// </para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="columnName"/> is null.</exception>
    /// <exception cref="ArgumentException">The column is not part of this table, or <paramref name="function"/> is <see cref="TotalsRowFunction.Custom"/> (use the other overload).</exception>
    /// <exception cref="InvalidOperationException"><see cref="HasTotalsRow"/> is false — call <see cref="AddTotalsRow"/> first.</exception>
    void SetColumnTotal(string columnName, TotalsRowFunction function);

    /// <summary>
    /// Configures the totals-row cell for <paramref name="columnName"/>
    /// to evaluate <paramref name="customFormula"/>. Sets the table-
    /// metadata function to <see cref="TotalsRowFunction.Custom"/> and
    /// writes the formula into the cell as-is. A leading <c>=</c> is
    /// optional and stripped.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException">The column is not part of this table, or the formula body is empty.</exception>
    /// <exception cref="InvalidOperationException"><see cref="HasTotalsRow"/> is false.</exception>
    void SetColumnTotal(string columnName, string customFormula);

    /// <summary>
    /// Sets the totals-row cell text for <paramref name="columnName"/>
    /// (decision I-64) — typically used for the leading "Total" label
    /// on the first column. The label takes precedence over any
    /// function configured on the same column.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException">The column is not part of this table.</exception>
    /// <exception cref="InvalidOperationException"><see cref="HasTotalsRow"/> is false.</exception>
    void SetColumnTotalLabel(string columnName, string label);

    /// <summary>
    /// The table-style name (e.g. <c>"TableStyleMedium2"</c>), or
    /// <c>null</c> when no style is applied. See <see cref="TableStyles"/>
    /// for common constants; any Excel-recognized style name is accepted.
    /// </summary>
    string? StyleName { get; set; }

    /// <summary>
    /// Escape hatch — direct access to the underlying Open XML SDK
    /// <see cref="DocumentFormat.OpenXml.Packaging.TableDefinitionPart"/>
    /// (I-82). The table's content lives in its own OPC part, so the hatch
    /// hands out the part (reach the DOM via <c>Table</c>). Same contract as
    /// <see cref="IWorkbook.Underlying"/>: direct mutation is supported but
    /// not synchronized with wrapper state.
    /// </summary>
    DocumentFormat.OpenXml.Packaging.TableDefinitionPart Underlying { get; }
}

/// <summary>
/// Common built-in table style names accepted by
/// <see cref="ISheet.AddTable"/> and <see cref="ITable.StyleName"/>.
/// Excel ships ~60 built-in styles; this is a curated subset of the
/// most-asked-for names. Any Excel-recognized style string works —
/// constants are convenience, not a closed set.
/// </summary>
public static class TableStyles
{
    /// <summary><c>"TableStyleLight1"</c> — light, neutral.</summary>
    public const string Light1 = "TableStyleLight1";
    /// <summary><c>"TableStyleLight9"</c> — light blue.</summary>
    public const string Light9 = "TableStyleLight9";
    /// <summary><c>"TableStyleLight15"</c> — light, plain.</summary>
    public const string Light15 = "TableStyleLight15";
    /// <summary><c>"TableStyleMedium2"</c> — Excel's "Blue, Table Style Medium 2" (the default in the UI).</summary>
    public const string Medium2 = "TableStyleMedium2";
    /// <summary><c>"TableStyleMedium9"</c> — strong blue.</summary>
    public const string Medium9 = "TableStyleMedium9";
    /// <summary><c>"TableStyleMedium16"</c> — orange.</summary>
    public const string Medium16 = "TableStyleMedium16";
    /// <summary><c>"TableStyleDark1"</c> — dark, neutral.</summary>
    public const string Dark1 = "TableStyleDark1";
    /// <summary><c>"TableStyleDark9"</c> — dark blue.</summary>
    public const string Dark9 = "TableStyleDark9";
}
