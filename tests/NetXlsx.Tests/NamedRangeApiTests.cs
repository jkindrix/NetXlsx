// Coverage for the v0.7 sub-slice B named-range API.

using System;
using System.IO;
using System.Linq;
using FluentAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class NamedRangeApiTests
{
    private static readonly string[] ExpectedNames = new[] { "Range1", "Range2", "Range3" };

    // ---- AddNamedRange — basic happy paths ----------------------------

    [Fact]
    public void AddNamedRange_Workbook_Scope_Roundtrips_Name_And_Formula()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");

        var n = wb.AddNamedRange("Sales", "Data!$A$1:$A$10");

        n.Name.Should().Be("Sales");
        n.Formula.Should().Be("Data!$A$1:$A$10");
        n.SheetScope.Should().BeNull();
    }

    [Fact]
    public void AddNamedRange_Strips_Optional_Leading_Equals()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        var n = wb.AddNamedRange("Sales", "=Data!$A$1:$A$10");
        n.Formula.Should().Be("Data!$A$1:$A$10");
    }

    [Fact]
    public void AddNamedRange_Sheet_Scoped_Reports_Sheet_Name()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        wb.AddSheet("Other");

        var n = wb.AddNamedRange("LocalRange", "Data!$A$1:$A$5", sheetScope: "Other");

        n.SheetScope.Should().Be("Other");
    }

    [Fact]
    public void NamedRanges_Returns_All_Defined_Ranges()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        wb.AddNamedRange("Range1", "Data!$A$1");
        wb.AddNamedRange("Range2", "Data!$A$2");
        wb.AddNamedRange("Range3", "Data!$A$3", sheetScope: "Data");

        wb.NamedRanges.Should().HaveCount(3);
        wb.NamedRanges.Select(r => r.Name).Should().BeEquivalentTo(ExpectedNames);
    }

    [Fact]
    public void NamedRanges_On_Fresh_Workbook_Is_Empty()
    {
        using var wb = Workbook.Create();
        wb.NamedRanges.Should().BeEmpty();
    }

    // ---- Argument validation -----------------------------------------

    [Fact]
    public void AddNamedRange_Null_Name_Throws_ArgumentNullException()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        ((Action)(() => wb.AddNamedRange(null!, "Data!$A$1"))).Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void AddNamedRange_Null_Formula_Throws_ArgumentNullException()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        ((Action)(() => wb.AddNamedRange("X", null!))).Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void AddNamedRange_Empty_Name_Or_Formula_Throws_ArgumentException()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        ((Action)(() => wb.AddNamedRange("", "Data!$A$1"))).Should().Throw<ArgumentException>();
        ((Action)(() => wb.AddNamedRange("X", ""))).Should().Throw<ArgumentException>();
    }

    [Fact]
    public void AddNamedRange_Unknown_SheetScope_Throws_SheetNameException()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        ((Action)(() => wb.AddNamedRange("X", "Data!$A$1", sheetScope: "DoesNotExist")))
            .Should().Throw<SheetNameException>();
    }

    // ---- Duplicate detection -----------------------------------------

    [Fact]
    public void AddNamedRange_Duplicate_At_Same_Workbook_Scope_Throws()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        wb.AddNamedRange("X", "Data!$A$1");
        ((Action)(() => wb.AddNamedRange("X", "Data!$A$2")))
            .Should().Throw<ArgumentException>()
            .WithMessage("*already exists*");
    }

    [Fact]
    public void AddNamedRange_Duplicate_Is_Case_Insensitive_At_Same_Scope()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        wb.AddNamedRange("Total", "Data!$A$1");
        ((Action)(() => wb.AddNamedRange("TOTAL", "Data!$A$2")))
            .Should().Throw<ArgumentException>();
    }

    [Fact]
    public void AddNamedRange_Same_Name_Different_Scope_Throws_For_NPOI_Compatibility()
    {
        // Excel itself permits workbook-scope and sheet-scope names to
        // coexist with the same text, but NPOI 2.7.x rejects this at
        // XSSFName.ValidateName. v1 requires workbook-wide unique names
        // regardless of scope. Documented as an NPOI constraint in
        // implementation-notes.md; revisit if/when NPOI relaxes this.
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        wb.AddNamedRange("X", "Data!$A$1");
        ((Action)(() => wb.AddNamedRange("X", "Data!$A$2", sheetScope: "Data")))
            .Should().Throw<ArgumentException>()
            .WithMessage("*unique workbook-wide*");
    }

    // ---- Round-trip through formulas ----------------------------------

    [Fact]
    public void Named_Range_Survives_Save_Open_And_Is_Usable_In_Formulas()
    {
        var path = Path.Combine(Path.GetTempPath(), $"named-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var data = wb.AddSheet("Data");
                data["A1"].SetNumber(10);
                data["A2"].SetNumber(20);
                data["A3"].SetNumber(30);
                wb.AddNamedRange("Sales", "Data!$A$1:$A$3");

                var summary = wb.AddSheet("Summary");
                summary["A1"].SetFormula("=SUM(Sales)");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb.NamedRanges.Should().ContainSingle().Which.Name.Should().Be("Sales");
                wb["Summary"]["A1"].GetFormula().Should().Be("=SUM(Sales)");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
