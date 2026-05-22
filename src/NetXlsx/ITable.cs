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
    /// Whether the table carries a totals row. Read-only in v1.1 —
    /// adding totals requires per-column totals-row functions, which
    /// is a v1.2 surface. Reach through <see cref="Underlying"/> for
    /// the full NPOI surface.
    /// </summary>
    bool HasTotalsRow { get; }

    /// <summary>
    /// The table-style name (e.g. <c>"TableStyleMedium2"</c>), or
    /// <c>null</c> when no style is applied. See <see cref="TableStyles"/>
    /// for common constants; any Excel-recognized style name is accepted.
    /// </summary>
    string? StyleName { get; set; }

    /// <summary>
    /// Escape hatch — direct access to the underlying NPOI <c>XSSFTable</c>.
    /// Same contract as <see cref="IWorkbook.Underlying"/>: direct
    /// mutation is supported but not synchronized with wrapper state.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFTable Underlying { get; }
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
