// Coverage for the v0.3.x IRow surface and the new sheet[r,c] indexer.

using System;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class RowApiTests
{
    [Fact]
    public void AppendRow_Starts_At_Row_1_For_Empty_Sheet()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var row = sheet.AppendRow();
        row.Index.Should().Be(1);
    }

    [Fact]
    public void AppendRow_Continues_After_LastRowNum_On_Used_Sheet()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("a");
        sheet["A5"].SetString("b");   // gap doesn't matter — appends after max
        var row = sheet.AppendRow();
        row.Index.Should().Be(6);
    }

    [Fact]
    public void Row_Auto_Materializes_Missing_Row()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var row = sheet.Row(10);
        row.Index.Should().Be(10);
        row.Cell(1).Kind.Should().Be(CellKind.Empty);
    }

    [Fact]
    public void Row_Out_Of_Range_Throws()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        Action zero = () => sheet.Row(0);
        zero.Should().Throw<ArgumentOutOfRangeException>();
        Action overflow = () => sheet.Row(CellAddress.MaxRow + 1);
        overflow.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Sheet_Indexer_By_Row_Col_Is_One_Based_And_Matches_A1_Form()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet[1, 1].SetString("from r,c");
        sheet["A1"].GetString().Should().Be("from r,c");

        sheet[10, 27].SetString("AA10");
        sheet["AA10"].GetString().Should().Be("AA10");
    }

    [Fact]
    public void Sheet_Indexer_Row_Col_Out_Of_Range_Throws()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        Action rowZero = () => { var _ = sheet[0, 1]; };
        rowZero.Should().Throw<ArgumentOutOfRangeException>();
        Action colZero = () => { var _ = sheet[1, 0]; };
        colZero.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Row_Indexer_By_Column_Letter_Works()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var row = sheet.AppendRow();
        row["B"].SetString("col B");
        row["AA"].SetString("col AA");
        sheet["B1"].GetString().Should().Be("col B");
        sheet["AA1"].GetString().Should().Be("col AA");
    }

    [Fact]
    public void Row_Set_Is_Fluent_And_Writes_All_Scalar_Kinds()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.AppendRow()
            .Set(1, "str")
            .Set(2, 1.5)
            .Set(3, 2.25m)
            .Set(4, 42)
            .Set(5, 9_000_000_000L)
            .Set(6, true);

        sheet["A1"].GetString().Should().Be("str");
        sheet["B1"].GetNumber().Should().Be(1.5);
        sheet["C1"].GetNumber().Should().Be(2.25);
        sheet["D1"].GetNumber().Should().Be(42.0);
        sheet["E1"].GetNumber().Should().Be(9_000_000_000.0);
        sheet["F1"].GetBool().Should().Be(true);
    }

    [Fact]
    public void Cell_SetNumber_Int_And_Long_Resolve_Unambiguously()
    {
        // The literal-42 ambiguity the recipe surfaced is now gone.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetNumber(42);          // resolves to SetNumber(int)
        sheet["A2"].SetNumber(42L);         // resolves to SetNumber(long)
        sheet["A1"].GetNumber().Should().Be(42.0);
        sheet["A2"].GetNumber().Should().Be(42.0);
    }

    [Fact]
    public void Row_Index_And_Sheet_Backreference_Are_Correct()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var row = sheet.AppendRow();
        row.Sheet.Should().BeSameAs(sheet);
        row.Index.Should().Be(1);

        var row2 = sheet.AppendRow();
        row2.Index.Should().Be(2);
    }
}
