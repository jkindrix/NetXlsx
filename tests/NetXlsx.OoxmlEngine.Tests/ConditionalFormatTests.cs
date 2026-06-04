// I-82 engine swap — CF/validation/tables/autofilter/sort slice: conditional
// formatting conformance.
//
// Mirrors the NPOI engine's ConditionalFormatTests contract on the Open XML
// SDK engine (decision I-73): every factory family lands as the right
// <cfRule> (cellIs/expression/colorScale + operator + formulas), one
// <conditionalFormatting sqref> per AddConditionalFormatting call, count /
// remove semantics, argument validation, and a file round-trip. cfRule
// styles land as deduped dxf entries in styles.xml (the NPOI engine
// allocates one per rule; the dedup is a deliberate pool-discipline
// improvement, invisible through the public API).
//
// The priority test pins the deliberate divergence from NPOI: priorities are
// allocated max+1 across the sheet, so a remove-then-add can never mint the
// duplicate priorities NPOI's count+1 emits (observed in the oracle dump).
//
// CF has no public read-back beyond ConditionalFormattingCount, so emission
// parity vs the NPOI engine is asserted by the projection test at the bottom
// (SDK-quirk #11 habit), normalizing known cosmetics (NPOI's nonconforming
// 6-digit rgb, its meaningless dxfId on colorScale rules, dxf id numbering).

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.OoxmlEngine.Tests;

public class ConditionalFormatTests
{
    private static S.Worksheet WorksheetOf(IWorkbook wb)
        => wb.Underlying.WorkbookPart!.WorksheetParts.Single().Worksheet!;

    private static S.ConditionalFormattingRule SingleRuleOf(IWorkbook wb)
        => WorksheetOf(wb).Elements<S.ConditionalFormatting>().Single()
            .Elements<S.ConditionalFormattingRule>().Single();

    [Fact]
    public void CellValueGreaterThan_Builds_CellIs_Rule()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.CellValueGreaterThan("50", new CellStyle { Bold = true }));

        s.ConditionalFormattingCount.Should().Be(1);
        var rule = SingleRuleOf(wb);
        rule.Type!.InnerText.Should().Be("cellIs");
        rule.Operator!.InnerText.Should().Be("greaterThan");
        rule.Priority!.Value.Should().Be(1);
        rule.FormatId.Should().NotBeNull("the bold style must land as a dxf");
        rule.Elements<S.Formula>().Single().Text.Should().Be("50");
    }

    [Fact]
    public void CellValueBetween_Builds_Two_Formulas()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.CellValueBetween("10", "90", new CellStyle { Italic = true }));

        var rule = SingleRuleOf(wb);
        rule.Operator!.InnerText.Should().Be("between");
        rule.Elements<S.Formula>().Select(f => f.Text).Should().Equal("10", "90");
    }

    [Theory]
    [InlineData("Equal", "equal")]
    [InlineData("NotEqual", "notEqual")]
    [InlineData("GreaterThanOrEqual", "greaterThanOrEqual")]
    [InlineData("LessThanOrEqual", "lessThanOrEqual")]
    [InlineData("LessThan", "lessThan")]
    public void CellValue_Comparisons_Map_Operators(string kind, string expectedOp)
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        var style = new CellStyle { Bold = true };
        var rule = kind switch
        {
            "Equal" => ConditionalFormat.CellValueEqual("5", style),
            "NotEqual" => ConditionalFormat.CellValueNotEqual("5", style),
            "GreaterThanOrEqual" => ConditionalFormat.CellValueGreaterThanOrEqual("5", style),
            "LessThanOrEqual" => ConditionalFormat.CellValueLessThanOrEqual("5", style),
            "LessThan" => ConditionalFormat.CellValueLessThan("5", style),
            _ => throw new ArgumentException("bad kind"),
        };
        s.AddConditionalFormatting("A1:A5", rule);

        SingleRuleOf(wb).Operator!.InnerText.Should().Be(expectedOp);
    }

    [Fact]
    public void Formula_Rule_Builds_Expression_Type()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.Formula("=$A1>50", new CellStyle { Bold = true }));

        var rule = SingleRuleOf(wb);
        rule.Type!.InnerText.Should().Be("expression");
        rule.Operator.Should().BeNull();
        // The leading '=' is stripped by the factory.
        rule.Elements<S.Formula>().Single().Text.Should().Be("$A1>50");
    }

    [Fact]
    public void ColorScale_Two_Color_Builds_Min_Max_Stops()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.ColorScale(Color.FromRgb(255, 0, 0), Color.FromRgb(0, 255, 0)));

        var rule = SingleRuleOf(wb);
        rule.Type!.InnerText.Should().Be("colorScale");
        rule.FormatId.Should().BeNull("a colorScale carries no dxf (NPOI's dxfId=0 there is a meaningless artifact)");
        var scale = rule.GetFirstChild<S.ColorScale>()!;
        scale.Elements<S.ConditionalFormatValueObject>().Select(v => v.Type!.InnerText)
            .Should().Equal("min", "max");
        scale.Elements<S.Color>().Select(c => c.Rgb!.Value)
            .Should().Equal("FFFF0000", "FF00FF00");
    }

    [Fact]
    public void ColorScale_Three_Color_Adds_Percentile_Stop()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.ColorScale(
                Color.FromRgb(0xF8, 0x69, 0x6B),
                Color.FromRgb(0xFF, 0xEB, 0x84),
                Color.FromRgb(0x63, 0xBE, 0x7B)));

        var scale = SingleRuleOf(wb).GetFirstChild<S.ColorScale>()!;
        var stops = scale.Elements<S.ConditionalFormatValueObject>().ToList();
        stops.Select(v => v.Type!.InnerText).Should().Equal("min", "percentile", "max");
        stops[1].Val!.Value.Should().Be("50");
        scale.Elements<S.Color>().Should().HaveCount(3);
    }

    [Fact]
    public void Multiple_Rules_On_Same_Range_Share_One_Element()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.CellValueGreaterThan("90", new CellStyle { Bold = true }),
            ConditionalFormat.CellValueLessThan("10", new CellStyle { Italic = true }));

        s.ConditionalFormattingCount.Should().Be(1);
        var cf = WorksheetOf(wb).Elements<S.ConditionalFormatting>().Single();
        cf.SequenceOfReferences!.InnerText.Should().Be("A1:A10");
        cf.Elements<S.ConditionalFormattingRule>().Select(r => r.Priority!.Value)
            .Should().Equal(1, 2);
    }

    [Fact]
    public void Multiple_AddConditionalFormatting_Calls_Append_In_Order()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A5",
            ConditionalFormat.CellValueGreaterThan("50", new CellStyle { Bold = true }));
        s.AddConditionalFormatting("B1:B5",
            ConditionalFormat.CellValueLessThan("10", new CellStyle { Italic = true }));

        s.ConditionalFormattingCount.Should().Be(2);
        var cfs = WorksheetOf(wb).Elements<S.ConditionalFormatting>().ToList();
        cfs.Select(c => c.SequenceOfReferences!.InnerText).Should().Equal("A1:A5", "B1:B5");
        // Priorities continue across calls.
        cfs[1].Elements<S.ConditionalFormattingRule>().Single().Priority!.Value.Should().Be(2);
    }

    [Fact]
    public void RemoveConditionalFormatting_Removes_The_Indexed_Element()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A5",
            ConditionalFormat.CellValueGreaterThan("50", new CellStyle { Bold = true }));
        s.AddConditionalFormatting("B1:B5",
            ConditionalFormat.CellValueLessThan("10", new CellStyle { Italic = true }));

        s.RemoveConditionalFormatting(0);

        s.ConditionalFormattingCount.Should().Be(1);
        WorksheetOf(wb).Elements<S.ConditionalFormatting>().Single()
            .SequenceOfReferences!.InnerText.Should().Be("B1:B5");
    }

    [Fact]
    public void Remove_Then_Add_Never_Duplicates_Priorities()
    {
        // Deliberate divergence from NPOI (whose count+1 allocation mints
        // duplicate priorities after a removal): priorities are max+1.
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A5", ConditionalFormat.CellValueGreaterThan("1", new CellStyle { Bold = true }));
        s.AddConditionalFormatting("B1:B5", ConditionalFormat.CellValueGreaterThan("2", new CellStyle { Bold = true }));
        s.AddConditionalFormatting("C1:C5", ConditionalFormat.CellValueGreaterThan("3", new CellStyle { Bold = true }));
        s.RemoveConditionalFormatting(1);   // drops priority 2
        s.AddConditionalFormatting("D1:D5", ConditionalFormat.CellValueGreaterThan("4", new CellStyle { Bold = true }));

        var priorities = WorksheetOf(wb).Descendants<S.ConditionalFormattingRule>()
            .Select(r => r.Priority!.Value).ToList();
        priorities.Should().OnlyHaveUniqueItems();
        priorities.Should().Equal(1, 3, 4);
    }

    [Fact]
    public void Rejects_Null_Range()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        Action act = () => s.AddConditionalFormatting(null!,
            ConditionalFormat.CellValueGreaterThan("1", new CellStyle()));
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Rejects_Empty_Rules()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        Action act = () => s.AddConditionalFormatting("A1:A5");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void CellValue_Rule_Survives_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.CreateOoxml())
        {
            var s = wb.AddSheet("S");
            for (int i = 1; i <= 5; i++) s[$"A{i}"].SetNumber(i * 20);
            s.AddConditionalFormatting("A1:A5",
                ConditionalFormat.CellValueGreaterThan("50", new CellStyle { Bold = true }));
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.OpenOoxml(ms);
        opened["S"].ConditionalFormattingCount.Should().Be(1);
        SingleRuleOf(opened).Elements<S.Formula>().Single().Text.Should().Be("50");
    }

    // ---- dxf (differential format) content + dedup --------------------------

    private static S.DifferentialFormats? DxfsOf(IWorkbook wb)
        => wb.Underlying.WorkbookPart!.WorkbookStylesPart!
            .Stylesheet!.GetFirstChild<S.DifferentialFormats>();

    [Fact]
    public void Style_Lands_As_A_Dxf_With_Font_And_Fill()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A5", ConditionalFormat.CellValueGreaterThan("1",
            new CellStyle { Bold = true, Italic = true, Background = Color.FromRgb(0xFF, 0xC7, 0xCE) }));

        var dxf = DxfsOf(wb)!.Elements<S.DifferentialFormat>().Single();
        var font = dxf.GetFirstChild<S.Font>()!;
        font.GetFirstChild<S.Bold>().Should().NotBeNull();
        font.GetFirstChild<S.Italic>().Should().NotBeNull();
        var fill = dxf.GetFirstChild<S.Fill>()!.GetFirstChild<S.PatternFill>()!;
        fill.PatternType!.InnerText.Should().Be("solid");
        fill.ForegroundColor!.Rgb!.Value.Should().Be("FFFFC7CE");
        fill.BackgroundColor!.Indexed!.Value.Should().Be(64u);
    }

    [Fact]
    public void Identical_Styles_Dedup_To_One_Dxf()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        var red = new CellStyle { Background = Color.FromRgb(0xFF, 0xC7, 0xCE) };
        s.AddConditionalFormatting("A1:A5", ConditionalFormat.CellValueGreaterThan("1", red));
        s.AddConditionalFormatting("B1:B5", ConditionalFormat.CellValueLessThan("9", red));

        var dxfs = DxfsOf(wb)!;
        dxfs.Elements<S.DifferentialFormat>().Should().HaveCount(1);
        dxfs.Count!.Value.Should().Be(1u);
        WorksheetOf(wb).Descendants<S.ConditionalFormattingRule>()
            .Select(r => r.FormatId!.Value).Should().Equal(0u, 0u);
    }

    // ---- emission projection pins (CF has no public read-back beyond the
    // count, so the persisted XML is the observable). Was the cross-engine
    // equivalence test; at the v2.0.0 cutover the NPOI half was collapsed
    // onto these LITERALS — captured from the green pre-flip state where
    // both engines projected identically (A1 disposition (b)). --------------

    private sealed record CfObs(string Sqref, string Type, string? Op, string[] Formulas,
        bool? DxfBold, bool? DxfItalic, string? DxfFillRgb, string[] ScaleStops, string[] ScaleColors);

    private static readonly string[] s_f30 = { "30" };
    private static readonly string[] s_f10To20 = { "10", "20" };
    private static readonly string[] s_fIsNumber = { "ISNUMBER(B1)" };
    private static readonly string[] s_scaleStops = { "min", "percentile", "max" };
    private static readonly string[] s_scaleColors = { "F8696B", "FFEB84", "63BE7B" };

    [Fact]
    public void Cf_Emission_Matches_The_Pinned_Projection()
    {
        static string NormalizeRgb(string? rgb)
            => rgb is null ? "" : (rgb.Length == 8 ? rgb[2..] : rgb);

        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-cf-par-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddSheet("S");
                for (int r = 1; r <= 5; r++) s[r, 1].SetNumber(r * 10);
                s.AddConditionalFormatting("A1:A5",
                    ConditionalFormat.CellValueGreaterThan("30", new CellStyle { Background = Color.FromRgb(0xFF, 0xC7, 0xCE) }),
                    ConditionalFormat.CellValueBetween("10", "20", new CellStyle { Bold = true, Italic = true }));
                s.AddConditionalFormatting("B1:B5",
                    ConditionalFormat.Formula("ISNUMBER(B1)", new CellStyle { Bold = true }));
                s.AddConditionalFormatting("C1:C5",
                    ConditionalFormat.ColorScale(
                        Color.FromRgb(0xF8, 0x69, 0x6B), Color.FromRgb(0xFF, 0xEB, 0x84), Color.FromRgb(0x63, 0xBE, 0x7B)));
                wb.Save(path);
            }
            using var opened = Workbook.Open(path);
            var ws = opened.Underlying.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var dxfs = opened.Underlying.WorkbookPart!.WorkbookStylesPart!
                .Stylesheet!.GetFirstChild<S.DifferentialFormats>()
                ?.Elements<S.DifferentialFormat>().ToList();

            var projected = ws.Elements<S.ConditionalFormatting>().SelectMany(cf =>
                cf.Elements<S.ConditionalFormattingRule>().Select(rule =>
                {
                    var type = rule.Type!.InnerText!;
                    // Resolve the dxf SHAPE rather than comparing ids (the
                    // engine dedups; colorScale rules carry no meaningful dxf).
                    S.DifferentialFormat? dxf =
                        type != "colorScale" && rule.FormatId?.Value is uint id && dxfs is not null && id < dxfs.Count
                            ? dxfs[(int)id] : null;
                    var font = dxf?.GetFirstChild<S.Font>();
                    var fill = dxf?.GetFirstChild<S.Fill>()?.GetFirstChild<S.PatternFill>();
                    var scale = rule.GetFirstChild<S.ColorScale>();
                    // Effective flag value: <b/> (val defaults true), omitted
                    // element (false), or explicit <i val="0"/> — compare the
                    // resolved boolean.
                    static bool EffectiveFlag(DocumentFormat.OpenXml.Spreadsheet.BooleanPropertyType? flag)
                        => flag is not null && (flag.Val?.Value ?? true);
                    return new CfObs(
                        cf.SequenceOfReferences!.InnerText!,
                        type,
                        rule.Operator?.InnerText,
                        rule.Elements<S.Formula>().Select(f => f.Text).ToArray(),
                        font is null ? null : EffectiveFlag(font.GetFirstChild<S.Bold>()),
                        font is null ? null : EffectiveFlag(font.GetFirstChild<S.Italic>()),
                        fill is null ? null : NormalizeRgb(fill.ForegroundColor?.Rgb?.Value),
                        scale?.Elements<S.ConditionalFormatValueObject>().Select(v => v.Type!.InnerText!).ToArray() ?? Array.Empty<string>(),
                        scale?.Elements<S.Color>().Select(c => NormalizeRgb(c.Rgb?.Value)).ToArray() ?? Array.Empty<string>());
                })).ToArray();

            projected.Should().BeEquivalentTo(new[]
            {
                new CfObs("A1:A5", "cellIs", "greaterThan", s_f30, null, null, "FFC7CE",
                    Array.Empty<string>(), Array.Empty<string>()),
                new CfObs("A1:A5", "cellIs", "between", s_f10To20, true, true, null,
                    Array.Empty<string>(), Array.Empty<string>()),
                new CfObs("B1:B5", "expression", null, s_fIsNumber, true, false, null,
                    Array.Empty<string>(), Array.Empty<string>()),
                new CfObs("C1:C5", "colorScale", null, Array.Empty<string>(), null, null, null,
                    s_scaleStops, s_scaleColors),
            }, o => o.WithStrictOrdering());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
