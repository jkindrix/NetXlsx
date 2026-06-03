// Data validation per design §6.4.4 / decision I-55.
// A sealed class with static factories for the common Excel
// validation types. Applied to a range via ISheet.AddValidation.

using System;
using NPOI.SS.UserModel;

namespace NetXlsx;

/// <summary>
/// A data-validation rule that can be applied to a range of cells via
/// <see cref="ISheet.AddValidation"/>. Construct via the static
/// factories — direct construction is not supported.
/// <para>
/// v1.1 covers Excel's most-used constraint families: list (explicit
/// values or formula reference), integer / decimal comparisons, date
/// range, text length, and custom formula. Less common constraints
/// (time of day, specific operator combinations) reach through the
/// NPOI <c>IDataValidationHelper</c> exposed by
/// <see cref="ISheet.Underlying"/>.
/// </para>
/// </summary>
public sealed class DataValidation
{
    // Engine-agnostic descriptor (I-82 engine swap): the Open XML SDK engine
    // cannot call the NPOI-typed Build closure, so each factory also captures
    // the OOXML CT_DataValidation axes (type / operator / formula1 / formula2)
    // directly. The NPOI engine keeps using Build; the SDK engine reads these.

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

    private readonly Func<IDataValidationHelper, IDataValidationConstraint> _build;

    internal ConstraintKind Kind { get; }
    internal CompareOp Operator { get; }
    internal string Formula1 { get; }
    internal string? Formula2 { get; }

    private DataValidation(
        Func<IDataValidationHelper, IDataValidationConstraint> build,
        ConstraintKind kind, CompareOp op, string formula1, string? formula2 = null)
    {
        _build = build;
        Kind = kind;
        Operator = op;
        Formula1 = formula1;
        Formula2 = formula2;
    }

    /// <summary>Internal: materialize this rule against a sheet's helper.</summary>
    internal IDataValidationConstraint Build(IDataValidationHelper helper) => _build(helper);

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
        return new DataValidation(
            h => h.CreateExplicitListConstraint(values),
            // OOXML encodes an explicit list as a quoted comma-joined formula1
            // ("Red,Green,Blue") — the same encoding NPOI's constraint emits.
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
        return new DataValidation(
            h => h.CreateFormulaListConstraint(formula),
            ConstraintKind.List, CompareOp.None, formula);
    }

    // ---- Integer ------------------------------------------------------

    /// <summary>Accepts integer values in <c>[min, max]</c> (inclusive).</summary>
    public static DataValidation IntegerBetween(int min, int max) =>
        Numeric(ValidationType.INTEGER, OperatorType.BETWEEN, ConstraintKind.Whole, CompareOp.Between, min.ToString(System.Globalization.CultureInfo.InvariantCulture), max.ToString(System.Globalization.CultureInfo.InvariantCulture));

    /// <summary>Accepts only integers equal to <paramref name="value"/>.</summary>
    public static DataValidation IntegerEqual(int value) =>
        Numeric(ValidationType.INTEGER, OperatorType.EQUAL, ConstraintKind.Whole, CompareOp.Equal, value.ToString(System.Globalization.CultureInfo.InvariantCulture), null);

    /// <summary>Accepts integers strictly greater than <paramref name="value"/>.</summary>
    public static DataValidation IntegerGreaterThan(int value) =>
        Numeric(ValidationType.INTEGER, OperatorType.GREATER_THAN, ConstraintKind.Whole, CompareOp.GreaterThan, value.ToString(System.Globalization.CultureInfo.InvariantCulture), null);

    /// <summary>Accepts integers strictly less than <paramref name="value"/>.</summary>
    public static DataValidation IntegerLessThan(int value) =>
        Numeric(ValidationType.INTEGER, OperatorType.LESS_THAN, ConstraintKind.Whole, CompareOp.LessThan, value.ToString(System.Globalization.CultureInfo.InvariantCulture), null);

    // ---- Decimal ------------------------------------------------------

    /// <summary>Accepts numeric values in <c>[min, max]</c> (inclusive).</summary>
    public static DataValidation DecimalBetween(double min, double max) =>
        Numeric(ValidationType.DECIMAL, OperatorType.BETWEEN, ConstraintKind.Decimal, CompareOp.Between,
            min.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            max.ToString("R", System.Globalization.CultureInfo.InvariantCulture));

    // ---- Date ---------------------------------------------------------

    /// <summary>
    /// Accepts date values in <c>[start, end]</c> (inclusive). Constructs the
    /// underlying constraint with the Excel <c>DATE(yyyy,m,d)</c> formula form
    /// so the validation survives a round-trip without locale interference.
    /// </summary>
    public static DataValidation DateBetween(DateOnly start, DateOnly end) =>
        new(h => h.CreateDateConstraint(
            OperatorType.BETWEEN,
            $"DATE({start.Year},{start.Month},{start.Day})",
            $"DATE({end.Year},{end.Month},{end.Day})",
            null),
            ConstraintKind.Date, CompareOp.Between,
            $"DATE({start.Year},{start.Month},{start.Day})",
            $"DATE({end.Year},{end.Month},{end.Day})");

    // ---- Text length --------------------------------------------------

    /// <summary>Accepts strings whose length is &lt;= <paramref name="max"/>.</summary>
    public static DataValidation TextLengthAtMost(int max) =>
        new(h => h.CreateTextLengthConstraint(
            OperatorType.LESS_OR_EQUAL,
            max.ToString(System.Globalization.CultureInfo.InvariantCulture),
            null),
            ConstraintKind.TextLength, CompareOp.LessThanOrEqual,
            max.ToString(System.Globalization.CultureInfo.InvariantCulture));

    /// <summary>Accepts strings whose length is &gt;= <paramref name="min"/>.</summary>
    public static DataValidation TextLengthAtLeast(int min) =>
        new(h => h.CreateTextLengthConstraint(
            OperatorType.GREATER_OR_EQUAL,
            min.ToString(System.Globalization.CultureInfo.InvariantCulture),
            null),
            ConstraintKind.TextLength, CompareOp.GreaterThanOrEqual,
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
        return new DataValidation(
            h => h.CreateCustomConstraint(formula),
            ConstraintKind.Custom, CompareOp.None, formula);
    }

    // ---- Internal -----------------------------------------------------

    private static DataValidation Numeric(int validationType, int op, ConstraintKind kind, CompareOp cmp, string f1, string? f2) =>
        new(h => h.CreateNumericConstraint(validationType, op, f1, f2), kind, cmp, f1, f2);
}
