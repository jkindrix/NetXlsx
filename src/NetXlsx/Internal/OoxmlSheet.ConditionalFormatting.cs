// I-82 engine swap — Open XML SDK-backed conditional formatting
// (CF/validation/tables/autofilter/sort slice). Surface per decision I-73.
//
// OOXML stores each AddConditionalFormatting call as ONE
// <conditionalFormatting sqref="…"> worksheet child holding one <cfRule> per
// rule. <conditionalFormatting> is a 0..* member of CT_Worksheet's strict
// child sequence — the long-carried slice-7 carry-forward: GetOrInsert's
// get-or-insert-singleton shape would wrongly return an existing element, so
// the insert goes through OoxmlSchemaOrder.Insert (the always-create
// companion; a new element lands after existing same-rank siblings).
//
// A cfRule's style is a dxf (differential format) in styles.xml referenced by
// @dxfId — built by OoxmlStylePool.GetOrCreateDifferentialFormat, which
// models exactly the axes the NPOI engine's ApplyStyle honors (bold / italic
// / background fill).
//
// Oracle-checked against the NPOI engine's emitted XML (SDK-quirk #11 habit):
// cellIs/expression/colorScale types, 1-2 <formula> children, and the
// min/percentile-50/max cfvo triple for 3-color scales all match. Deliberate
// divergences from NPOI cosmetics: @aboveAverage="1" is the schema default
// (omitted); NPOI's colorScale rules carry a meaningless dxfId="0" artifact
// (omitted); colorScale colors are written 8-digit ARGB rather than NPOI's
// nonconforming 6-digit form; and rule @priority is allocated max+1 across
// the sheet rather than NPOI's count+1, which can emit AMBIGUOUS DUPLICATE
// priorities after a removal (observed in the oracle dump — two rules with
// priority="4").

using System;
using System.Linq;
using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    public void AddConditionalFormatting(string a1Range, params ConditionalFormat[] rules)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(a1Range);
        ArgumentNullException.ThrowIfNull(rules);
        if (rules.Length == 0) throw new ArgumentException("At least one rule is required.", nameof(rules));

        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        var cf = new S.ConditionalFormatting
        {
            SequenceOfReferences = new ListValue<StringValue>
            {
                InnerText = CellAddress.FormatRange(r1, c1, r2, c2),
            },
        };

        int priority = NextRulePriority();
        foreach (var rule in rules)
            cf.AppendChild(BuildRule(rule, priority++));

        OoxmlSchemaOrder.Insert(Worksheet, cf);
    }

    public int ConditionalFormattingCount
    {
        get
        {
            ThrowIfUnusable();
            return Worksheet.Elements<S.ConditionalFormatting>().Count();
        }
    }

    public void RemoveConditionalFormatting(int index)
    {
        ThrowIfUnusable();
        // ElementAt throws ArgumentOutOfRangeException for a bad index — the
        // NPOI engine's list removal fails loud on a bad index too. Any dxf
        // the removed rules referenced stays in styles.xml (NPOI parity;
        // orphaned dxfs are harmless).
        Worksheet.Elements<S.ConditionalFormatting>().ElementAt(index).Remove();
    }

    // Rule priorities are unique-ascending across the whole sheet; lower wins.
    // max+1 (not NPOI's count+1) so a remove-then-add can never mint a
    // duplicate priority.
    private int NextRulePriority()
    {
        int max = 0;
        foreach (var rule in Worksheet.Elements<S.ConditionalFormatting>()
                     .SelectMany(cf => cf.Elements<S.ConditionalFormattingRule>()))
        {
            if (rule.Priority?.Value is int p && p > max) max = p;
        }
        return max + 1;
    }

    private S.ConditionalFormattingRule BuildRule(ConditionalFormat rule, int priority)
    {
        var cfRule = new S.ConditionalFormattingRule { Priority = priority };
        switch (rule.Kind)
        {
            case ConditionalFormat.RuleKind.CellValue:
                cfRule.Type = S.ConditionalFormatValues.CellIs;
                cfRule.Operator = ToCfOperator(rule.OperatorName!);
                if (rule.Style is not null)
                    cfRule.FormatId = _workbook.StylePool.GetOrCreateDifferentialFormat(rule.Style);
                cfRule.AppendChild(new S.Formula(rule.Formula1!));
                if (rule.Formula2 is not null)
                    cfRule.AppendChild(new S.Formula(rule.Formula2));
                break;

            case ConditionalFormat.RuleKind.Formula:
                cfRule.Type = S.ConditionalFormatValues.Expression;
                if (rule.Style is not null)
                    cfRule.FormatId = _workbook.StylePool.GetOrCreateDifferentialFormat(rule.Style);
                cfRule.AppendChild(new S.Formula(rule.Formula1!));
                break;

            default: // ColorScale
                cfRule.Type = S.ConditionalFormatValues.ColorScale;
                cfRule.AppendChild(BuildColorScale(rule));
                break;
        }
        return cfRule;
    }

    // <colorScale> wants its cfvo children first, then one <color> per cfvo:
    // min / [percentile 50] / max — the same stops NPOI's two- and three-color
    // scale rules produce.
    private static S.ColorScale BuildColorScale(ConditionalFormat rule)
    {
        var scale = new S.ColorScale();
        scale.AppendChild(new S.ConditionalFormatValueObject { Type = S.ConditionalFormatValueObjectValues.Min });
        if (rule.MidColor is not null)
            scale.AppendChild(new S.ConditionalFormatValueObject
            {
                Type = S.ConditionalFormatValueObjectValues.Percentile,
                Val = "50",
            });
        scale.AppendChild(new S.ConditionalFormatValueObject { Type = S.ConditionalFormatValueObjectValues.Max });

        scale.AppendChild(ToCfColor(rule.MinColor!.Value));
        if (rule.MidColor is not null) scale.AppendChild(ToCfColor(rule.MidColor.Value));
        scale.AppendChild(ToCfColor(rule.MaxColor!.Value));
        return scale;
    }

    private static S.Color ToCfColor(Color c) => new() { Rgb = $"FF{c.R:X2}{c.G:X2}{c.B:X2}" };

    private static S.ConditionalFormattingOperatorValues ToCfOperator(string name) => name switch
    {
        "between" => S.ConditionalFormattingOperatorValues.Between,
        "notBetween" => S.ConditionalFormattingOperatorValues.NotBetween,
        "equal" => S.ConditionalFormattingOperatorValues.Equal,
        "notEqual" => S.ConditionalFormattingOperatorValues.NotEqual,
        "greaterThan" => S.ConditionalFormattingOperatorValues.GreaterThan,
        "lessThan" => S.ConditionalFormattingOperatorValues.LessThan,
        "greaterThanOrEqual" => S.ConditionalFormattingOperatorValues.GreaterThanOrEqual,
        _ => S.ConditionalFormattingOperatorValues.LessThanOrEqual,
    };
}
