// Frozen v1.0 set of built-in Excel number format strings per design
// §6.11 / decision I12. Additions in later releases require an explicit
// decision row.

namespace NetXlsx;

/// <summary>
/// Curated set of Excel number-format strings for common cases. These
/// are pass-through Excel format codes (§7.2) — NetXlsx does not
/// localize them, and Excel renders them per its own culture at display
/// time. Use as the value of <c>CellStyle.NumberFormat</c> or as the
/// argument to <see cref="ICell.NumberFormat(string)"/>.
/// </summary>
public static class NumberFormats
{
    /// <summary>Excel's default format. Renders numbers as-is, text as-is.</summary>
    public const string General = "General";

    /// <summary>Text format. Renders any value as literal text.</summary>
    public const string Text = "@";

    /// <summary>Integer with no thousands separator.</summary>
    public const string Integer = "0";

    /// <summary>Integer with thousands separator.</summary>
    public const string Number = "#,##0";

    /// <summary>Two-decimal number with thousands separator.</summary>
    public const string NumberTwo = "#,##0.00";

    /// <summary>Scientific notation, 2 decimal places.</summary>
    public const string Scientific = "0.00E+00";

    /// <summary>Whole-percent (e.g. <c>50%</c>).</summary>
    public const string Percent = "0%";

    /// <summary>Two-decimal percent (e.g. <c>50.25%</c>).</summary>
    public const string PercentTwo = "0.00%";

    /// <summary>USD currency with $ symbol and two decimals.</summary>
    public const string Currency = "$#,##0.00";

    /// <summary>Currency with thousands separator but no symbol.</summary>
    public const string CurrencyNoSymbol = "#,##0.00";

    /// <summary>Accountant-style: negatives in red.</summary>
    public const string Accounting = "$#,##0.00;[Red]-$#,##0.00";

    /// <summary>ISO-shaped date.</summary>
    public const string Date = "yyyy-mm-dd";

    /// <summary>ISO-shaped date + time.</summary>
    public const string DateTime = "yyyy-mm-dd hh:mm:ss";

    /// <summary>Time of day, hours:minutes:seconds.</summary>
    public const string Time = "hh:mm:ss";

    /// <summary>Elapsed time — supports values greater than 24h (§7.9).</summary>
    public const string Duration = "[h]:mm:ss";
}
