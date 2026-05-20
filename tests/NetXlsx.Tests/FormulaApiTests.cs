// Coverage for the v0.7 sub-slice A formula API (ICell.SetFormula / GetFormula).

using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class FormulaApiTests
{
    // ---- SetFormula ---------------------------------------------------

    [Fact]
    public void SetFormula_With_Leading_Equals_Stores_Body_Without_It()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["C1"].SetFormula("=A1+B1");
        sheet["C1"].Kind.Should().Be(CellKind.Formula);
        sheet["C1"].GetFormula().Should().Be("=A1+B1");
    }

    [Fact]
    public void SetFormula_Without_Leading_Equals_Is_Accepted()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["C1"].SetFormula("A1+B1");
        sheet["C1"].GetFormula().Should().Be("=A1+B1");
    }

    [Fact]
    public void SetFormula_Null_Throws_ArgumentNullException()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].SetFormula(null!)))
            .Should().Throw<ArgumentNullException>();
    }

    [Theory]
    [InlineData("")]
    [InlineData("=")]
    public void SetFormula_Empty_Body_Throws_FormulaException(string input)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].SetFormula(input)))
            .Should().Throw<FormulaException>()
            .WithMessage("*empty*");
    }

    [Fact]
    public void SetFormula_Unparseable_Throws_FormulaException_With_Original_Text()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].SetFormula("=SUM(")))
            .Should().Throw<FormulaException>()
            .WithMessage("*SUM(*");
    }

    // ---- GetFormula ---------------------------------------------------

    [Fact]
    public void GetFormula_On_Non_Formula_Cell_Returns_Null()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("hello");
        sheet["A1"].GetFormula().Should().BeNull();
        sheet["A2"].GetFormula().Should().BeNull(); // empty cell
    }

    [Fact]
    public void GetFormula_Roundtrips_Sheet_Qualified_Reference()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        var sheet = wb.AddSheet("Summary");
        sheet["A1"].SetFormula("=SUM(Data!A1:A10)");
        sheet["A1"].GetFormula().Should().Be("=SUM(Data!A1:A10)");
    }

    // ---- No pre-computation (design §7.8) -----------------------------

    [Fact]
    public void SetFormula_Does_Not_Pre_Compute_Cached_Value()
    {
        // Per decision #46 / §7.8: cached value must be left for Excel
        // to recompute. NPOI exposes the cached result via
        // XSSFCell.NumericCellValue (or similar); when the formula has
        // not been evaluated, NPOI's CachedFormulaResultType is Numeric
        // with value 0 (its default), not the computed sum. Easier
        // proof: directly check the cached result type is the spec'd
        // default and the workbook is not flagged force-recalc.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetNumber(2);
        sheet["B1"].SetNumber(3);
        sheet["C1"].SetFormula("=A1+B1");

        // No evaluator was called — NPOI's cached result remains its
        // default. Excel will recompute on open. We assert through the
        // escape hatch because there's no public surface for "what
        // value did we pre-compute" (we don't).
        var raw = sheet["C1"].Underlying;
        raw.CachedFormulaResultType.Should().Be(NPOI.SS.UserModel.CellType.Numeric);
        raw.NumericCellValue.Should().Be(0.0,
            "design decision #46 — formulas are written with no cached value; Excel recalculates on open");
    }

    // ---- Replace semantics --------------------------------------------

    [Fact]
    public void SetFormula_After_SetNumber_Replaces_The_Cell()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetNumber(42);
        sheet["A1"].Kind.Should().Be(CellKind.Number);

        sheet["A1"].SetFormula("=1+1");
        sheet["A1"].Kind.Should().Be(CellKind.Formula);
        sheet["A1"].GetFormula().Should().Be("=1+1");
    }

    [Fact]
    public void Clear_After_SetFormula_Resets_Kind_To_Empty()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetFormula("=1+1");
        sheet["A1"].Clear();
        sheet["A1"].Kind.Should().Be(CellKind.Empty);
        sheet["A1"].GetFormula().Should().BeNull();
    }

    // ---- Round-trip ---------------------------------------------------

    [Fact]
    public void Formula_Roundtrips_Through_Save_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"formula-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("S");
                sheet["A1"].SetNumber(10);
                sheet["A2"].SetNumber(20);
                sheet["A3"].SetFormula("=SUM(A1:A2)");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var cell = wb["S"]["A3"];
                cell.Kind.Should().Be(CellKind.Formula);
                cell.GetFormula().Should().Be("=SUM(A1:A2)");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
