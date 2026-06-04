using System;

namespace NetXlsx;

/// <summary>
/// Describes a conditional formatting rule for
/// <see cref="ISheet.AddConditionalFormatting"/> (decision I-73).
/// Constructed via static factories — no public constructor.
/// </summary>
public sealed class ConditionalFormat
{
    internal enum RuleKind { CellValue, Formula, ColorScale }

    /// <summary>Internal: the comparison operator for CellValue rules (I-82 engine-agnostic descriptor).</summary>
    internal enum CompareOp
    {
        Between,
        NotBetween,
        Equal,
        NotEqual,
        GreaterThan,
        LessThan,
        GreaterThanOrEqual,
        LessThanOrEqual,
    }

    internal RuleKind Kind { get; }
    internal CompareOp? Operator { get; }
    internal string? Formula1 { get; }
    internal string? Formula2 { get; }
    internal CellStyle? Style { get; }
    internal Color? MinColor { get; }
    internal Color? MidColor { get; }
    internal Color? MaxColor { get; }

    /// <summary>
    /// Internal: the OOXML <c>ST_ConditionalFormattingOperator</c> name for
    /// <see cref="Operator"/>, or <c>null</c> when no operator applies.
    /// </summary>
    internal string? OperatorName => Operator switch
    {
        CompareOp.Between => "between",
        CompareOp.NotBetween => "notBetween",
        CompareOp.Equal => "equal",
        CompareOp.NotEqual => "notEqual",
        CompareOp.GreaterThan => "greaterThan",
        CompareOp.LessThan => "lessThan",
        CompareOp.GreaterThanOrEqual => "greaterThanOrEqual",
        CompareOp.LessThanOrEqual => "lessThanOrEqual",
        _ => null,
    };

    private ConditionalFormat(RuleKind kind, CompareOp? op = null,
        string? formula1 = null, string? formula2 = null, CellStyle? style = null,
        Color? minColor = null, Color? midColor = null, Color? maxColor = null)
    {
        Kind = kind;
        Operator = op;
        Formula1 = formula1;
        Formula2 = formula2;
        Style = style;
        MinColor = minColor;
        MidColor = midColor;
        MaxColor = maxColor;
    }

    /// <summary>Highlight cells whose value is greater than <paramref name="value"/>.</summary>
    public static ConditionalFormat CellValueGreaterThan(string value, CellStyle style) =>
        CellValueRule(CompareOp.GreaterThan, value, null, style);

    /// <summary>Highlight cells whose value is less than <paramref name="value"/>.</summary>
    public static ConditionalFormat CellValueLessThan(string value, CellStyle style) =>
        CellValueRule(CompareOp.LessThan, value, null, style);

    /// <summary>Highlight cells whose value is between <paramref name="min"/> and <paramref name="max"/> (inclusive).</summary>
    public static ConditionalFormat CellValueBetween(string min, string max, CellStyle style) =>
        CellValueRule(CompareOp.Between, min, max, style);

    /// <summary>Highlight cells whose value equals <paramref name="value"/>.</summary>
    public static ConditionalFormat CellValueEqual(string value, CellStyle style) =>
        CellValueRule(CompareOp.Equal, value, null, style);

    /// <summary>Highlight cells whose value does not equal <paramref name="value"/>.</summary>
    public static ConditionalFormat CellValueNotEqual(string value, CellStyle style) =>
        CellValueRule(CompareOp.NotEqual, value, null, style);

    /// <summary>Highlight cells whose value is greater than or equal to <paramref name="value"/>.</summary>
    public static ConditionalFormat CellValueGreaterThanOrEqual(string value, CellStyle style) =>
        CellValueRule(CompareOp.GreaterThanOrEqual, value, null, style);

    /// <summary>Highlight cells whose value is less than or equal to <paramref name="value"/>.</summary>
    public static ConditionalFormat CellValueLessThanOrEqual(string value, CellStyle style) =>
        CellValueRule(CompareOp.LessThanOrEqual, value, null, style);

    /// <summary>Highlight cells where <paramref name="formula"/> evaluates to <c>TRUE</c>.</summary>
    public static ConditionalFormat Formula(string formula, CellStyle style)
    {
        ArgumentNullException.ThrowIfNull(formula);
        ArgumentNullException.ThrowIfNull(style);
        var body = formula.StartsWith('=') ? formula.Substring(1) : formula;
        return new ConditionalFormat(RuleKind.Formula, formula1: body, style: style);
    }

    /// <summary>Apply a two-color gradient scale across the range.</summary>
    public static ConditionalFormat ColorScale(Color min, Color max) =>
        new(RuleKind.ColorScale, minColor: min, maxColor: max);

    /// <summary>Apply a three-color gradient scale across the range.</summary>
    public static ConditionalFormat ColorScale(Color min, Color mid, Color max) =>
        new(RuleKind.ColorScale, minColor: min, midColor: mid, maxColor: max);

    private static ConditionalFormat CellValueRule(CompareOp op, string value1, string? value2, CellStyle style)
    {
        ArgumentNullException.ThrowIfNull(value1);
        ArgumentNullException.ThrowIfNull(style);
        return new ConditionalFormat(RuleKind.CellValue, op: op, formula1: value1, formula2: value2, style: style);
    }
}
