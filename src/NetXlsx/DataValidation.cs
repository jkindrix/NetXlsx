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
    private readonly Func<IDataValidationHelper, IDataValidationConstraint> _build;

    private DataValidation(Func<IDataValidationHelper, IDataValidationConstraint> build)
    {
        _build = build;
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
        return new DataValidation(h => h.CreateExplicitListConstraint(values));
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
        return new DataValidation(h => h.CreateFormulaListConstraint(formula));
    }

    // ---- Integer ------------------------------------------------------

    /// <summary>Accepts integer values in <c>[min, max]</c> (inclusive).</summary>
    public static DataValidation IntegerBetween(int min, int max) =>
        Numeric(ValidationType.INTEGER, OperatorType.BETWEEN, min.ToString(System.Globalization.CultureInfo.InvariantCulture), max.ToString(System.Globalization.CultureInfo.InvariantCulture));

    /// <summary>Accepts only integers equal to <paramref name="value"/>.</summary>
    public static DataValidation IntegerEqual(int value) =>
        Numeric(ValidationType.INTEGER, OperatorType.EQUAL, value.ToString(System.Globalization.CultureInfo.InvariantCulture), null);

    /// <summary>Accepts integers strictly greater than <paramref name="value"/>.</summary>
    public static DataValidation IntegerGreaterThan(int value) =>
        Numeric(ValidationType.INTEGER, OperatorType.GREATER_THAN, value.ToString(System.Globalization.CultureInfo.InvariantCulture), null);

    /// <summary>Accepts integers strictly less than <paramref name="value"/>.</summary>
    public static DataValidation IntegerLessThan(int value) =>
        Numeric(ValidationType.INTEGER, OperatorType.LESS_THAN, value.ToString(System.Globalization.CultureInfo.InvariantCulture), null);

    // ---- Decimal ------------------------------------------------------

    /// <summary>Accepts numeric values in <c>[min, max]</c> (inclusive).</summary>
    public static DataValidation DecimalBetween(double min, double max) =>
        Numeric(ValidationType.DECIMAL, OperatorType.BETWEEN,
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
            null));

    // ---- Text length --------------------------------------------------

    /// <summary>Accepts strings whose length is &lt;= <paramref name="max"/>.</summary>
    public static DataValidation TextLengthAtMost(int max) =>
        new(h => h.CreateTextLengthConstraint(
            OperatorType.LESS_OR_EQUAL,
            max.ToString(System.Globalization.CultureInfo.InvariantCulture),
            null));

    /// <summary>Accepts strings whose length is &gt;= <paramref name="min"/>.</summary>
    public static DataValidation TextLengthAtLeast(int min) =>
        new(h => h.CreateTextLengthConstraint(
            OperatorType.GREATER_OR_EQUAL,
            min.ToString(System.Globalization.CultureInfo.InvariantCulture),
            null));

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
        return new DataValidation(h => h.CreateCustomConstraint(formula));
    }

    // ---- Internal -----------------------------------------------------

    private static DataValidation Numeric(int validationType, int op, string f1, string? f2) =>
        new(h => h.CreateNumericConstraint(validationType, op, f1, f2));
}
