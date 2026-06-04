// I-82 engine swap — formulas (formulas/comments/hyperlinks slice) conformance.
//
// Mirrors the NPOI-engine FormulaApiTests behavioral contract on the Open XML
// SDK engine: '=' normalization, empty-body / structurally-broken input ->
// FormulaException, Kind classification, replace semantics, Clear, and the
// Save/Open round-trip. Adds DOM-level assertions (via IWorkbook.Underlying)
// that no cached <v> is pre-computed (design decision #46 / §7.8).

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using S = DocumentFormat.OpenXml.Spreadsheet;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class FormulaTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-formula-{Guid.NewGuid():N}.xlsx");

    private static S.Cell CellElement(IWorkbook wb, string reference)
        => wb.Underlying.WorkbookPart!.WorksheetParts.Single()
            .Worksheet!.Descendants<S.Cell>().Single(c => c.CellReference?.Value == reference);

    // ---- SetFormula ---------------------------------------------------------

    [Fact]
    public void SetFormula_With_Leading_Equals_Stores_Body_Without_It()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["C1"].SetFormula("=A1+B1");
        sheet["C1"].Kind.Should().Be(CellKind.Formula);
        sheet["C1"].GetFormula().Should().Be("=A1+B1");
        CellElement(wb, "C1").CellFormula!.Text.Should().Be("A1+B1");
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

    // The SDK engine has no formula parser (NPOI does); it fails loud on the
    // structural breakage Excel would reject with a repair prompt. The message
    // carries the original text, matching the NPOI-engine contract.
    [Theory]
    [InlineData("=SUM(")]            // unbalanced '('
    [InlineData("=A1)+B1")]          // unbalanced ')'
    [InlineData("=IF(A1=\"x,1,2)")]  // unterminated string literal
    [InlineData("='Sheet1!A1")]      // unterminated quoted sheet name
    public void SetFormula_Structurally_Broken_Throws_FormulaException_With_Original_Text(string input)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].SetFormula(input)))
            .Should().Throw<FormulaException>()
            .WithMessage($"*{input.Substring(1)}*");
    }

    [Theory]
    [InlineData("=IF(A1=\"x(\",1,2)")]       // '(' inside a string literal is fine
    [InlineData("='Q1 (final)'!A1+1")]       // parens inside a quoted sheet name
    [InlineData("=\"a\"\"b\"&A1")]           // doubled-quote escape inside a literal
    public void SetFormula_Accepts_Literals_Containing_Parens_And_Escaped_Quotes(string input)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet["A1"].SetFormula(input))).Should().NotThrow();
        sheet["A1"].GetFormula().Should().Be(input);
    }

    // ---- GetFormula ---------------------------------------------------------

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

    // ---- No pre-computation (design §7.8 / decision #46) --------------------

    [Fact]
    public void SetFormula_Does_Not_Pre_Compute_Cached_Value()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetNumber(2);
        sheet["B1"].SetNumber(3);
        sheet["C1"].SetFormula("=A1+B1");

        var c = CellElement(wb, "C1");
        c.CellFormula.Should().NotBeNull();
        c.CellValue.Should().BeNull(
            "design decision #46 — formulas are written with no cached value; Excel recalculates on open");
    }

    // ---- Replace semantics --------------------------------------------------

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
        CellElement(wb, "A1").CellValue.Should().BeNull("the numeric <v> must not survive as a stale cached result");
    }

    [Fact]
    public void SetString_After_SetFormula_Removes_The_Formula()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetFormula("=1+1");
        sheet["A1"].SetString("plain");
        sheet["A1"].Kind.Should().Be(CellKind.String);
        sheet["A1"].GetFormula().Should().BeNull();
        sheet["A1"].GetString().Should().Be("plain");
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

    // ---- Round-trip ---------------------------------------------------------

    [Fact]
    public void Formula_Roundtrips_Through_Save_Open()
    {
        var path = TempXlsxPath();
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
