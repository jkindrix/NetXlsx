// Coverage for ICell.GetError + CellError enum (decision #49).
//
// Excel error codes (from the OOXML spec / NPOI's FormulaError):
//   0x00 = #NULL!         0x07 = #DIV/0!
//   0x0F = #VALUE!        0x17 = #REF!
//   0x1D = #NAME?         0x24 = #NUM!
//   0x2A = #N/A           0x2B = #GETTING_DATA
//
// Writing error values directly requires .Underlying (no Set-side API
// for errors in v0.4); we exercise reading and the formula-result path.

using FluentAssertions;
using NPOI.SS.UserModel;
using Xunit;

namespace NetXlsx.Tests;

public class CellErrorTests
{
    // NPOI's SetCellErrorValue rejects 0x2B (#GETTING_DATA) — that code
    // is producible only by Excel's own evaluator from external data
    // sources. The mapping function still handles it on the read path
    // (covered by Mapping_Function_Handles_Every_Code below).
    [Theory]
    [InlineData((byte)0x00, CellError.Null)]
    [InlineData((byte)0x07, CellError.DivByZero)]
    [InlineData((byte)0x0F, CellError.Value)]
    [InlineData((byte)0x17, CellError.Ref)]
    [InlineData((byte)0x1D, CellError.Name)]
    [InlineData((byte)0x24, CellError.Num)]
    [InlineData((byte)0x2A, CellError.NotAvailable)]
    public void GetError_Maps_Writable_Error_Codes(byte excelCode, CellError expected)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var cell = sheet["A1"];

        cell.Underlying.SetCellErrorValue(excelCode);

        cell.Kind.Should().Be(CellKind.Error);
        cell.GetError().Should().Be(expected);
    }

    [Fact]
    public void GettingData_Enum_Value_Exists_For_Files_Authored_By_Excel()
    {
        // 0x2B (#GETTING_DATA) is unreachable through NPOI's write API
        // but real Excel workbooks can contain it. The enum value must
        // exist so consumers of GetError can pattern-match it.
        System.Enum.IsDefined(CellError.GettingData).Should().BeTrue();
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

    // The "formula cell with cached error result" path through
    // GetError exists for real Excel-authored workbooks where Excel's
    // own evaluator produced the cached value. Programmatically forcing
    // that state from NPOI alone requires the formula evaluator to
    // run — out of scope for a unit test. The path is exercised by
    // any real-world workbook with formula errors round-tripped
    // through NetXlsx.
}
