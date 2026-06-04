// I-82 engine swap — CF/validation/tables/autofilter/sort slice: data
// validation conformance.
//
// Mirrors the NPOI engine's DataValidationTests contract on the Open XML SDK
// engine (decision I-55): every factory family lands as the right
// CT_DataValidation (@type/@operator/formula1/formula2/@sqref/@allowBlank),
// multiple rules coexist with @count synced, arguments are validated, and
// rules survive a file round-trip.
//
// ISheet has no validation read-back API, so the cross-engine differential
// harness (public-API observations) cannot cover this surface. The emission-
// parity test below stands in for it: the SAME rules are written through BOTH
// engines and the two files' <dataValidations> are compared as normalized
// projections (schema defaults applied), per the SDK-quirk #11 oracle habit.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.Tests.Engine;

public class DataValidationTests
{
    private static S.DataValidations? ValidationsOf(IWorkbook wb)
        => wb.Underlying.WorkbookPart!.WorksheetParts.Single()
            .Worksheet!.GetFirstChild<S.DataValidations>();

    [Fact]
    public void AddValidation_List_Encodes_Quoted_Joined_Formula()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("A2:A5", DataValidation.List("Red", "Green", "Blue"));

        var dv = ValidationsOf(wb)!.Elements<S.DataValidation>().Single();
        dv.Type!.InnerText.Should().Be("list");
        dv.Formula1!.Text.Should().Be("\"Red,Green,Blue\"");
        dv.SequenceOfReferences!.InnerText.Should().Be("A2:A5");
        dv.AllowBlank!.Value.Should().BeTrue();
    }

    [Fact]
    public void AddValidation_ListFromRange_Keeps_Raw_Formula()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("B2:B5", DataValidation.ListFromRange("$Z$1:$Z$9"));

        var dv = ValidationsOf(wb)!.Elements<S.DataValidation>().Single();
        dv.Type!.InnerText.Should().Be("list");
        dv.Formula1!.Text.Should().Be("$Z$1:$Z$9");
    }

    [Fact]
    public void AddValidation_IntegerBetween_Builds_Whole_Between()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("C2:C5", DataValidation.IntegerBetween(1, 10));

        var dv = ValidationsOf(wb)!.Elements<S.DataValidation>().Single();
        dv.Type!.InnerText.Should().Be("whole");
        dv.Operator!.InnerText.Should().Be("between");
        dv.Formula1!.Text.Should().Be("1");
        dv.Formula2!.Text.Should().Be("10");
    }

    [Theory]
    [InlineData("Equal", "equal", "7")]
    [InlineData("GreaterThan", "greaterThan", "0")]
    [InlineData("LessThan", "lessThan", "100")]
    public void AddValidation_Integer_Comparisons_Map_Operators(string kind, string expectedOp, string expectedF1)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        var rule = kind switch
        {
            "Equal" => DataValidation.IntegerEqual(7),
            "GreaterThan" => DataValidation.IntegerGreaterThan(0),
            "LessThan" => DataValidation.IntegerLessThan(100),
            _ => throw new ArgumentException("bad kind"),
        };
        sh.AddValidation("D2:D5", rule);

        var dv = ValidationsOf(wb)!.Elements<S.DataValidation>().Single();
        dv.Type!.InnerText.Should().Be("whole");
        dv.Operator!.InnerText.Should().Be(expectedOp);
        dv.Formula1!.Text.Should().Be(expectedF1);
        dv.Formula2.Should().BeNull();
    }

    [Fact]
    public void AddValidation_DecimalBetween_Builds_Decimal()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("G2:G5", DataValidation.DecimalBetween(0.5, 9.5));

        var dv = ValidationsOf(wb)!.Elements<S.DataValidation>().Single();
        dv.Type!.InnerText.Should().Be("decimal");
        dv.Formula1!.Text.Should().Be("0.5");
        dv.Formula2!.Text.Should().Be("9.5");
    }

    [Fact]
    public void DateBetween_Uses_DATE_Formula_For_Locale_Stability()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("H2:H5", DataValidation.DateBetween(new DateOnly(2024, 1, 1), new DateOnly(2024, 12, 31)));

        var dv = ValidationsOf(wb)!.Elements<S.DataValidation>().Single();
        dv.Type!.InnerText.Should().Be("date");
        dv.Formula1!.Text.Should().Be("DATE(2024,1,1)");
        dv.Formula2!.Text.Should().Be("DATE(2024,12,31)");
    }

    [Theory]
    [InlineData("AtMost", "lessThanOrEqual", "10")]
    [InlineData("AtLeast", "greaterThanOrEqual", "2")]
    public void AddValidation_TextLength_Maps_Operators(string kind, string expectedOp, string expectedF1)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        var rule = kind == "AtMost" ? DataValidation.TextLengthAtMost(10) : DataValidation.TextLengthAtLeast(2);
        sh.AddValidation("I2:I5", rule);

        var dv = ValidationsOf(wb)!.Elements<S.DataValidation>().Single();
        dv.Type!.InnerText.Should().Be("textLength");
        dv.Operator!.InnerText.Should().Be(expectedOp);
        dv.Formula1!.Text.Should().Be(expectedF1);
    }

    [Fact]
    public void AddValidation_Custom_Keeps_Raw_Formula()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("K2:K5", DataValidation.Custom("ISNUMBER(K2)"));

        var dv = ValidationsOf(wb)!.Elements<S.DataValidation>().Single();
        dv.Type!.InnerText.Should().Be("custom");
        dv.Operator.Should().BeNull("operator is meaningless for custom; the schema default suffices");
        dv.Formula1!.Text.Should().Be("ISNUMBER(K2)");
    }

    [Fact]
    public void AddValidation_Multiple_Coexist_With_Count_Synced()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("A2:A5", DataValidation.List("a", "b"));
        sh.AddValidation("B2:B5", DataValidation.IntegerBetween(1, 10));
        sh.AddValidation("C2:C5", DataValidation.DecimalBetween(0, 1));
        sh.AddValidation("D2:D5", DataValidation.TextLengthAtMost(5));
        sh.AddValidation("E2:E5", DataValidation.Custom("ISNUMBER(E2)"));

        var container = ValidationsOf(wb)!;
        container.Elements<S.DataValidation>().Should().HaveCount(5);
        container.Count!.Value.Should().Be(5u);
        // Still a single container.
        wb.Underlying.WorkbookPart!.WorksheetParts.Single()
            .Worksheet!.Elements<S.DataValidations>().Should().HaveCount(1);
    }

    [Fact]
    public void AddValidation_Rejects_Null_Range_Or_Validation()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action nullRange = () => sh.AddValidation(null!, DataValidation.List("x"));
        Action nullRule = () => sh.AddValidation("A1:A2", null!);
        nullRange.Should().Throw<ArgumentNullException>();
        nullRule.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void AddValidation_Rejects_Invalid_Range()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.AddValidation("NOT_A_RANGE", DataValidation.List("x"));
        act.Should().Throw<InvalidCellAddressException>();
    }

    [Fact]
    public void AddValidation_Accepts_Single_Cell_Address()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("B2", DataValidation.IntegerBetween(1, 5));

        ValidationsOf(wb)!.Elements<S.DataValidation>().Single()
            .SequenceOfReferences!.InnerText.Should().Be("B2");
    }

    [Fact]
    public void Validation_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-dv-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh.AddValidation("A2:A5", DataValidation.List("Red", "Green"));
                sh.AddValidation("B2:B5", DataValidation.IntegerBetween(1, 10));
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var container = ValidationsOf(wb)!;
                container.Elements<S.DataValidation>().Should().HaveCount(2);
                container.Elements<S.DataValidation>().First().Formula1!.Text.Should().Be("\"Red,Green\"");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- emission projection pins (no public read-back API exists for
    // validations, so the persisted XML is the observable). Was the
    // cross-engine equivalence test; at the v2.0.0 cutover the NPOI half was
    // collapsed onto these LITERALS — captured from the green pre-flip state
    // where both engines emitted them identically (A1 disposition (b)). ----

    private sealed record DvObs(string Type, string Op, bool AllowBlank, string Sqref, string? F1, string? F2);

    [Fact]
    public void Validation_Emission_Matches_The_Pinned_Projection()
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-dv-par-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddSheet("S");
                s["A1"].SetString("h");
                s.AddValidation("A2:A5", DataValidation.List("Red", "Green", "Blue"));
                s.AddValidation("B2:B5", DataValidation.ListFromRange("$Z$1:$Z$9"));
                s.AddValidation("C2:C5", DataValidation.IntegerBetween(1, 10));
                s.AddValidation("D2", DataValidation.IntegerEqual(7));
                s.AddValidation("E2:E5", DataValidation.DecimalBetween(0.5, 9.5));
                s.AddValidation("F2:F5", DataValidation.DateBetween(new DateOnly(2024, 1, 1), new DateOnly(2024, 12, 31)));
                s.AddValidation("G2:G5", DataValidation.TextLengthAtMost(10));
                s.AddValidation("H2:H5", DataValidation.Custom("ISNUMBER(H2)"));
                wb.Save(path);
            }
            // Read the file back through the DOM and project with schema
            // defaults applied (absent @operator means between).
            using var opened = Workbook.Open(path);
            var projected = ValidationsOf(opened)!.Elements<S.DataValidation>()
                .Select(dv => new DvObs(
                    dv.Type!.InnerText!,
                    dv.Operator?.InnerText ?? "between",
                    dv.AllowBlank?.Value ?? false,
                    dv.SequenceOfReferences!.InnerText!,
                    dv.Formula1?.Text,
                    dv.Formula2?.Text))
                .ToArray();

            projected.Should().BeEquivalentTo(new[]
            {
                new DvObs("list", "between", true, "A2:A5", "\"Red,Green,Blue\"", null),
                new DvObs("list", "between", true, "B2:B5", "$Z$1:$Z$9", null),
                new DvObs("whole", "between", true, "C2:C5", "1", "10"),
                new DvObs("whole", "equal", true, "D2", "7", null),
                new DvObs("decimal", "between", true, "E2:E5", "0.5", "9.5"),
                new DvObs("date", "between", true, "F2:F5", "DATE(2024,1,1)", "DATE(2024,12,31)"),
                new DvObs("textLength", "lessThanOrEqual", true, "G2:G5", "10", null),
                new DvObs("custom", "between", true, "H2:H5", "ISNUMBER(H2)", null),
            }, o => o.WithStrictOrdering());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
