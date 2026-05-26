using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

/// <summary>
/// Describes a conditional formatting rule for
/// <see cref="ISheet.AddConditionalFormatting"/> (decision I-73).
/// Constructed via static factories — no public constructor.
/// </summary>
public sealed class ConditionalFormat
{
    internal enum RuleKind { CellValue, Formula, ColorScale }

    internal RuleKind Kind { get; }
    internal ComparisonOperator? Operator { get; }
    internal string? Formula1 { get; }
    internal string? Formula2 { get; }
    internal CellStyle? Style { get; }
    internal Color? MinColor { get; }
    internal Color? MidColor { get; }
    internal Color? MaxColor { get; }

    private ConditionalFormat(RuleKind kind, ComparisonOperator? op = null,
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
        CellValueRule(ComparisonOperator.GreaterThan, value, null, style);

    /// <summary>Highlight cells whose value is less than <paramref name="value"/>.</summary>
    public static ConditionalFormat CellValueLessThan(string value, CellStyle style) =>
        CellValueRule(ComparisonOperator.LessThan, value, null, style);

    /// <summary>Highlight cells whose value is between <paramref name="min"/> and <paramref name="max"/> (inclusive).</summary>
    public static ConditionalFormat CellValueBetween(string min, string max, CellStyle style) =>
        CellValueRule(ComparisonOperator.Between, min, max, style);

    /// <summary>Highlight cells whose value equals <paramref name="value"/>.</summary>
    public static ConditionalFormat CellValueEqual(string value, CellStyle style) =>
        CellValueRule(ComparisonOperator.Equal, value, null, style);

    /// <summary>Highlight cells whose value does not equal <paramref name="value"/>.</summary>
    public static ConditionalFormat CellValueNotEqual(string value, CellStyle style) =>
        CellValueRule(ComparisonOperator.NotEqual, value, null, style);

    /// <summary>Highlight cells whose value is greater than or equal to <paramref name="value"/>.</summary>
    public static ConditionalFormat CellValueGreaterThanOrEqual(string value, CellStyle style) =>
        CellValueRule(ComparisonOperator.GreaterThanOrEqual, value, null, style);

    /// <summary>Highlight cells whose value is less than or equal to <paramref name="value"/>.</summary>
    public static ConditionalFormat CellValueLessThanOrEqual(string value, CellStyle style) =>
        CellValueRule(ComparisonOperator.LessThanOrEqual, value, null, style);

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

    private static ConditionalFormat CellValueRule(ComparisonOperator op, string value1, string? value2, CellStyle style)
    {
        ArgumentNullException.ThrowIfNull(value1);
        ArgumentNullException.ThrowIfNull(style);
        return new ConditionalFormat(RuleKind.CellValue, op: op, formula1: value1, formula2: value2, style: style);
    }

    internal IConditionalFormattingRule CreateNpoiRule(ISheetConditionalFormatting scf)
    {
        IConditionalFormattingRule rule;
        switch (Kind)
        {
            case RuleKind.CellValue:
                rule = Formula2 != null
                    ? scf.CreateConditionalFormattingRule(Operator!.Value, Formula1!, Formula2)
                    : scf.CreateConditionalFormattingRule(Operator!.Value, Formula1!);
                ApplyStyle(rule);
                return rule;
            case RuleKind.Formula:
                rule = scf.CreateConditionalFormattingRule(Formula1!);
                ApplyStyle(rule);
                return rule;
            case RuleKind.ColorScale:
                rule = ((XSSFSheetConditionalFormatting)scf).CreateConditionalFormattingColorScaleRule();
                var cs = rule.ColorScaleFormatting;
                if (cs != null)
                {
                    var colors = cs.Colors;
                    if (MidColor is not null && colors.Length >= 3)
                    {
                        SetExtendedColor(colors[0], MinColor!.Value);
                        SetExtendedColor(colors[1], MidColor.Value);
                        SetExtendedColor(colors[2], MaxColor!.Value);
                    }
                    else if (colors.Length >= 2)
                    {
                        SetExtendedColor(colors[0], MinColor!.Value);
                        SetExtendedColor(colors[colors.Length - 1], MaxColor!.Value);
                    }
                }
                return rule;
            default:
                throw new InvalidOperationException($"Unknown rule kind: {Kind}");
        }
    }

    private void ApplyStyle(IConditionalFormattingRule rule)
    {
        if (Style is null) return;

        if (Style.Bold == true || Style.Italic == true)
        {
            var ff = rule.CreateFontFormatting();
            ff.SetFontStyle(Style.Bold == true, Style.Italic == true);
        }

        if (Style.Background is { } bg)
        {
            var pf = rule.CreatePatternFormatting();
            pf.FillBackgroundColor = IndexedColors.Automatic.Index;
            pf.FillPattern = FillPattern.SolidForeground;
            if (pf is XSSFPatternFormatting xpf)
            {
                var bgColor = new XSSFColor(new byte[] { bg.R, bg.G, bg.B });
                xpf.FillForegroundColorColor = bgColor;
            }
        }
    }

    private static void SetExtendedColor(IColor color, Color c)
    {
        if (color is XSSFColor xc)
            xc.SetRgb(new byte[] { c.R, c.G, c.B });
    }
}
