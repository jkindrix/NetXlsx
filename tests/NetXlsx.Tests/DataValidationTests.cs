// Coverage for the v1.1 data-validation slice: DataValidation factory
// methods, ISheet.AddValidation, file round-trip preservation.

using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class DataValidationTests
{
    // ---- Factory validation -------------------------------------------

    [Fact]
    public void List_Rejects_Empty_Values()
    {
        Action act = () => DataValidation.List(Array.Empty<string>());
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void List_Rejects_Null_Values()
    {
        Action act = () => DataValidation.List((string[])null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void ListFromRange_Rejects_Empty_Formula()
    {
        Action act = () => DataValidation.ListFromRange("");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Custom_Rejects_Empty_Formula()
    {
        Action act = () => DataValidation.Custom("");
        act.Should().Throw<ArgumentException>();
    }

    // ---- Apply to range ------------------------------------------------

    [Fact]
    public void AddValidation_List_Adds_To_Sheet()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("A1:A5", DataValidation.List("Yes", "No", "Maybe"));
        sh.Underlying.GetDataValidations().Count.Should().Be(1);
    }

    [Fact]
    public void AddValidation_IntegerBetween_Adds_To_Sheet()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("B2:B10", DataValidation.IntegerBetween(1, 100));
        sh.Underlying.GetDataValidations().Count.Should().Be(1);
    }

    [Fact]
    public void AddValidation_Multiple_Coexist()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("A1:A5", DataValidation.List("X", "Y"));
        sh.AddValidation("B1:B5", DataValidation.IntegerBetween(0, 10));
        sh.AddValidation("C1:C5", DataValidation.DateBetween(new DateOnly(2020, 1, 1), new DateOnly(2030, 12, 31)));
        sh.AddValidation("D1:D5", DataValidation.TextLengthAtMost(20));
        sh.AddValidation("E1:E5", DataValidation.Custom("ISNUMBER(E1)"));
        sh.Underlying.GetDataValidations().Count.Should().Be(5);
    }

    [Fact]
    public void AddValidation_Rejects_Null_Range_Or_Validation()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action a = () => sh.AddValidation(null!, DataValidation.IntegerEqual(5));
        a.Should().Throw<ArgumentNullException>();
        Action b = () => sh.AddValidation("A1", null!);
        b.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void AddValidation_Rejects_Invalid_Range()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.AddValidation("notarange", DataValidation.IntegerEqual(5));
        act.Should().Throw<InvalidCellAddressException>();
    }

    // ---- Single-cell range (no colon) ---------------------------------

    [Fact]
    public void AddValidation_Accepts_Single_Cell_Address()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("A1", DataValidation.IntegerBetween(1, 10));
        sh.Underlying.GetDataValidations().Count.Should().Be(1);
    }

    // ---- File round-trip ----------------------------------------------

    [Fact]
    public void Validation_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"dv-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh.AddValidation("A1:A10", DataValidation.List("Red", "Green", "Blue"));
                sh.AddValidation("B1:B10", DataValidation.IntegerBetween(0, 100));
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var validations = wb["S"].Underlying.GetDataValidations();
                validations.Count.Should().Be(2);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void DateBetween_Uses_DATE_Formula_For_Locale_Stability()
    {
        // Verify the constraint formula uses DATE(...) form rather than
        // a locale-specific date literal — important so the validation
        // survives a round-trip on machines with non-en-US date formats.
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddValidation("A1", DataValidation.DateBetween(new DateOnly(2026, 5, 22), new DateOnly(2027, 1, 1)));
        var dv = sh.Underlying.GetDataValidations()[0];
        dv.ValidationConstraint.Formula1.Should().Contain("DATE(2026,5,22)");
        dv.ValidationConstraint.Formula2.Should().Contain("DATE(2027,1,1)");
    }
}
