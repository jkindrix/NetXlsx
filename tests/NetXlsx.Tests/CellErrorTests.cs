// Coverage for ICell.GetError + CellError enum (decision #49).
//
// Excel error literals (OOXML stores the LITERAL in a t="e" cell's <v>):
//   #NULL!  #DIV/0!  #VALUE!  #REF!  #NAME?  #NUM!  #N/A  #GETTING_DATA
//
// Writing error values directly requires .Underlying (no Set-side API for
// errors); since v2.0.0 the hatch is the SDK <c> element, so the tests author
// the exact OOXML shape Excel itself writes — including #GETTING_DATA, which
// NPOI's old write API refused to produce.

using AwesomeAssertions;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.Tests;

public class CellErrorTests
{
    [Theory]
    [InlineData("#NULL!", CellError.Null)]
    [InlineData("#DIV/0!", CellError.DivByZero)]
    [InlineData("#VALUE!", CellError.Value)]
    [InlineData("#REF!", CellError.Ref)]
    [InlineData("#NAME?", CellError.Name)]
    [InlineData("#NUM!", CellError.Num)]
    [InlineData("#N/A", CellError.NotAvailable)]
    [InlineData("#GETTING_DATA", CellError.GettingData)]
    public void GetError_Maps_Every_Error_Literal(string literal, CellError expected)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var cell = sheet["A1"];

        // Author the error through the escape hatch — the t="e" + literal <v>
        // shape Excel writes.
        var c = cell.Underlying;
        c.DataType = S.CellValues.Error;
        c.CellValue = new S.CellValue(literal);

        cell.Kind.Should().Be(CellKind.Error);
        cell.GetError().Should().Be(expected);
    }

    [Fact]
    public void GetError_On_Non_Error_Cells_Returns_Null()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("hello");
        sheet["A2"].SetNumber(42);
        sheet["A3"].SetBool(true);

        sheet["A1"].GetError().Should().BeNull();
        sheet["A2"].GetError().Should().BeNull();
        sheet["A3"].GetError().Should().BeNull();
    }

    [Fact]
    public void GetError_Reads_A_Formula_Cells_Cached_Error()
    {
        // A formula whose cached result is an error is t="e" with both <f>
        // and the error literal in <v> — the shape Excel's evaluator writes.
        using var wb = Workbook.Create();
        var cell = wb.AddSheet("S")["A1"];
        cell.SetFormula("=1/0");
        var c = cell.Underlying;
        c.DataType = S.CellValues.Error;
        c.CellValue = new S.CellValue("#DIV/0!");

        cell.Kind.Should().Be(CellKind.Formula, "the formula classification wins");
        cell.GetError().Should().Be(CellError.DivByZero);
    }
}
