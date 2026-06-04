// Data validation per design §6.4.4 / decision I-55.
// A sealed class with static factories for the common Excel
// validation types. Applied to a range via ISheet.AddValidation.

using System;

namespace NetXlsx;

/// <summary>
/// A data-validation rule that can be applied to a range of cells via
/// <see cref="ISheet.AddValidation"/>. Construct via the static
/// factories — direct construction is not supported.
/// <para>
/// Covers Excel's most-used constraint families: list (explicit
/// values or formula reference), integer / decimal comparisons, date
/// range, text length, and custom formula. Less common constraints
/// (time of day, specific operator combinations) reach through the
/// worksheet DOM exposed by <see cref="ISheet.Underlying"/>.
/// </para>
/// </summary>
public sealed class DataValidation
{
    // Engine-agnostic descriptor (I-82): each factory captures the OOXML
    // CT_DataValidation axes (type / operator / formula1 / formula2)
    // directly; the engine emits the <dataValidation> element from them.

    /// <summary>Internal: the OOXML <c>@type</c> family.</summary>
    internal enum ConstraintKind
    {
        List,
        Whole,
        Decimal,
        Date,
        TextLength,
        Custom,
    }

    /// <summary>Internal: the OOXML <c>@operator</c>; None for list/custom.</summary>
    internal enum CompareOp
    {
        None,
        Between,
        Equal,
        GreaterThan,
        LessThan,
        LessThanOrEqual,
        GreaterThanOrEqual,
    }

    internal ConstraintKind Kind { get; }
    internal CompareOp Operator { get; }
    internal string Formula1 { get; }
    internal string? Formula2 { get; }

    private DataValidation(ConstraintKind kind, CompareOp op, string formula1, string? formula2 = null)
    {
        Kind = kind;
        Operator = op;
        Formula1 = formula1;
        Formula2 = formula2;
    }

    // ---- List (dropdown) ----------------------------------------------

    /// <summary>
    /// Dropdown picker with the literal values supplied. Excel limits
    /// the joined string to 255 characters — long enumerations should
    /// use <see cref="ListFromRange(string)"/> instead.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="values"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="values"/> is empty.</exception>
    public static DataValidation List(params string[] values)
    {
        ArgumentNullException.ThrowIfNull(values);
        if (values.Length == 0)
            throw new ArgumentException("List validation requires at least one value.", nameof(values));
        // OOXML encodes an explicit list as a quoted comma-joined formula1
        // ("Red,Green,Blue").
        return new DataValidation(
            ConstraintKind.List, CompareOp.None, "\"" + string.Join(",", values) + "\"");
    }

    /// <summary>
    /// Dropdown picker sourced from a formula reference (e.g.
    /// <c>"$A$1:$A$10"</c>, <c>"Sheet2!$A:$A"</c>, or a named-range
    /// reference). Use this for long enumerations that would exceed
    /// Excel's 255-character explicit-list limit.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="formula"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="formula"/> is empty.</exception>
    public static DataValidation ListFromRange(string formula)
    {
        ArgumentNullException.ThrowIfNull(formula);
        if (formula.Length == 0)
            throw new ArgumentException("Formula reference cannot be empty.", nameof(formula));
        return new DataValidation(ConstraintKind.List, CompareOp.None, formula);
    }

    // ---- Integer ------------------------------------------------------

    /// <summary>Accepts integer values in <c>[min, max]</c> (inclusive).</summary>
    public static DataValidation IntegerBetween(int min, int max) =>
        new(ConstraintKind.Whole, CompareOp.Between,
            min.ToString(System.Globalization.CultureInfo.InvariantCulture),
            max.ToString(System.Globalization.CultureInfo.InvariantCulture));

    /// <summary>Accepts only integers equal to <paramref name="value"/>.</summary>
    public static DataValidation IntegerEqual(int value) =>
        new(ConstraintKind.Whole, CompareOp.Equal,
            value.ToString(System.Globalization.CultureInfo.InvariantCulture));

    /// <summary>Accepts integers strictly greater than <paramref name="value"/>.</summary>
    public static DataValidation IntegerGreaterThan(int value) =>
        new(ConstraintKind.Whole, CompareOp.GreaterThan,
            value.ToString(System.Globalization.CultureInfo.InvariantCulture));

    /// <summary>Accepts integers strictly less than <paramref name="value"/>.</summary>
    public static DataValidation IntegerLessThan(int value) =>
        new(ConstraintKind.Whole, CompareOp.LessThan,
            value.ToString(System.Globalization.CultureInfo.InvariantCulture));

    // ---- Decimal ------------------------------------------------------

    /// <summary>Accepts numeric values in <c>[min, max]</c> (inclusive).</summary>
    public static DataValidation DecimalBetween(double min, double max) =>
        new(ConstraintKind.Decimal, CompareOp.Between,
            min.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            max.ToString("R", System.Globalization.CultureInfo.InvariantCulture));

    // ---- Date ---------------------------------------------------------

    /// <summary>
    /// Accepts date values in <c>[start, end]</c> (inclusive). Constructs the
    /// underlying constraint with the Excel <c>DATE(yyyy,m,d)</c> formula form
    /// so the validation survives a round-trip without locale interference.
    /// </summary>
    public static DataValidation DateBetween(DateOnly start, DateOnly end) =>
        new(ConstraintKind.Date, CompareOp.Between,
            $"DATE({start.Year},{start.Month},{start.Day})",
            $"DATE({end.Year},{end.Month},{end.Day})");

    // ---- Text length --------------------------------------------------

    /// <summary>Accepts strings whose length is &lt;= <paramref name="max"/>.</summary>
    public static DataValidation TextLengthAtMost(int max) =>
        new(ConstraintKind.TextLength, CompareOp.LessThanOrEqual,
            max.ToString(System.Globalization.CultureInfo.InvariantCulture));

    /// <summary>Accepts strings whose length is &gt;= <paramref name="min"/>.</summary>
    public static DataValidation TextLengthAtLeast(int min) =>
        new(ConstraintKind.TextLength, CompareOp.GreaterThanOrEqual,
            min.ToString(System.Globalization.CultureInfo.InvariantCulture));

    // ---- Custom -------------------------------------------------------

    /// <summary>
    /// Custom validation by Excel formula — the cell value is valid
    /// iff <paramref name="formula"/> evaluates to TRUE. Use
    /// relative references (e.g. <c>"ISNUMBER(A1)"</c>) for
    /// per-cell evaluation.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="formula"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="formula"/> is empty.</exception>
    public static DataValidation Custom(string formula)
    {
        ArgumentNullException.ThrowIfNull(formula);
        if (formula.Length == 0)
            throw new ArgumentException("Custom validation formula cannot be empty.", nameof(formula));
        return new DataValidation(ConstraintKind.Custom, CompareOp.None, formula);
    }
}
