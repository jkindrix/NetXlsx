// I-82 engine swap — cells & rows slice conformance.
//
// Value round-trips (string/number/bool), Kind, the materialization contract
// (decision #40), row/cell navigation + Set chaining, range Value/ClearContents
// + sparse vs dense enumeration, and the Excel ascending-order invariant for
// rows/cells. Also pins the deferred-surface boundary (SetDate etc. throw until
// their slice).

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class CellAndRowValueTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-cells-{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void Round_Trips_String_Number_Bool_Through_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddSheet("S");
                s["A1"].SetString("hello");
                s["B2"].SetNumber(42.5);
                s["C3"].SetBool(true);
                s["D4"].SetNumber(-1234567890123L);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var s = wb["S"];
                s["A1"].GetString().Should().Be("hello");
                s["A1"].Kind.Should().Be(CellKind.String);
                s["B2"].GetNumber().Should().Be(42.5);
                s["B2"].Kind.Should().Be(CellKind.Number);
                s["C3"].GetBool().Should().BeTrue();
                s["C3"].Kind.Should().Be(CellKind.Bool);
                s["D4"].GetNumber().Should().Be(-1234567890123d);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Never_Written_Cell_Is_Empty_And_Adds_No_Dom_Node()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("x");      // write one cell
        _ = s["Z99"].Kind;           // merely read another — must not materialize
        _ = s["Z99"].GetString();

        s["Z99"].Kind.Should().Be(CellKind.Empty);
        s["Z99"].GetString().Should().BeEmpty();
        s["Z99"].GetNumber().Should().BeNull();

        // Decision #40: reading must not add a <c> node to the DOM.
        var wsPart = wb.Underlying.WorkbookPart!.WorksheetParts.First();
        var refs = wsPart.Worksheet!.GetFirstChild<SheetData>()!
            .Descendants<Cell>().Select(c => c.CellReference!.Value).ToList();
        refs.Should().ContainSingle().Which.Should().Be("A1");
    }

    [Fact]
    public void GetString_Formats_Number_And_Bool()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetNumber(3.5);
        s["A2"].SetBool(false);
        s["A1"].GetString().Should().Be("3.5");
        s["A2"].GetString().Should().Be("FALSE");
    }

    [Fact]
    public void GetNumber_Returns_Bool_As_One_Or_Zero()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetBool(true);
        s["A2"].SetBool(false);
        s["A1"].GetNumber().Should().Be(1.0);
        s["A2"].GetNumber().Should().Be(0.0);
    }

    [Fact]
    public void Overwriting_A_Cell_Changes_Its_Kind()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetNumber(1);
        s["A1"].Kind.Should().Be(CellKind.Number);
        s["A1"].SetString("now text");
        s["A1"].Kind.Should().Be(CellKind.String);
        s["A1"].GetString().Should().Be("now text");
        s["A1"].GetNumber().Should().BeNull();
    }

    [Fact]
    public void Clear_Resets_Cell_To_Empty()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("x");
        s["A1"].Clear();
        s["A1"].Kind.Should().Be(CellKind.Empty);
        s["A1"].GetString().Should().BeEmpty();
    }

    [Fact]
    public void Indexer_RC_And_A1_Address_The_Same_Cell()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s[2, 3].SetString("here");          // row 2, col 3 == C2
        s["C2"].GetString().Should().Be("here");
        s["C2"].Address.Should().Be("C2");
        s[2, 3].RowIndex.Should().Be(2);
        s[2, 3].ColumnIndex.Should().Be(3);
    }

    [Fact]
    public void SetString_Preserves_Leading_And_Trailing_Whitespace()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S")["A1"].SetString("  pad  ");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
                wb["S"]["A1"].GetString().Should().Be("  pad  ");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Row_Set_Chaining_Writes_Across_Columns()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var row = s.AppendRow();
        row.Index.Should().Be(1);
        row.Set(1, "name").Set(2, 10).Set(3, true);
        s["A1"].GetString().Should().Be("name");
        s["B1"].GetNumber().Should().Be(10);
        s["C1"].GetBool().Should().BeTrue();
    }

    [Fact]
    public void AppendRow_Appends_After_Last_Written_Row()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A5"].SetString("five");           // highest written row is 5
        var appended = s.AppendRow();
        appended.Index.Should().Be(6);
    }

    [Fact]
    public void Range_Value_Sets_All_Then_ClearContents_Clears()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var range = s.Range("A1:B2");
        range.Address.Should().Be("A1:B2");
        range.Count.Should().Be(4);

        range.Value(7.0);
        range.EnumerateAll().Should().OnlyContain(c => c.GetNumber() == 7.0);

        range.ClearContents();
        range.EnumerateAll().Should().OnlyContain(c => c.Kind == CellKind.Empty);
    }

    [Fact]
    public void Range_Sparse_Enumeration_Yields_Only_Populated_Cells()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("a");
        s["C3"].SetString("c");             // B2 left empty

        // Sparse enumeration is row-major ascending, so an ordered Equal is the
        // stronger assertion (and avoids a constant-array argument / CA1861).
        var populated = s.Range("A1:C3").Select(c => c.Address).ToList();
        populated.Should().Equal("A1", "C3");

        s.Range("A1:C3").EnumerateAll().Count().Should().Be(9);  // dense
    }

    [Fact]
    public void Out_Of_Order_Writes_Persist_In_Ascending_Excel_Order()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        // Write deliberately out of order across rows and columns.
        s["C3"].SetString("c3");
        s["A1"].SetString("a1");
        s["B1"].SetString("b1");
        s["A3"].SetString("a3");

        var data = wb.Underlying.WorkbookPart!.WorksheetParts.First()
            .Worksheet!.GetFirstChild<SheetData>()!;

        var rowOrder = data.Elements<Row>().Select(r => (int)r.RowIndex!.Value).ToList();
        rowOrder.Should().BeInAscendingOrder().And.Equal(1, 3);

        var row1Cols = data.Elements<Row>().First(r => r.RowIndex!.Value == 1)
            .Elements<Cell>().Select(c => c.CellReference!.Value).ToList();
        row1Cols.Should().Equal("A1", "B1");   // ascending column order within the row
    }

    [Fact]
    public void Column_Exposes_Index_And_Letter()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.Column(28).Letter.Should().Be("AB");
        s.Column("AB").Index.Should().Be(28);
    }

    // (The cells & rows slice's Deferred_Formula_Setter_Throws_NotImplemented
    // guard retired when SetFormula landed — FormulaTests carries its coverage.)

    [Fact]
    public void SetString_Over_Length_Limit_Throws()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var tooLong = new string('x', 32768);   // Excel cell text cap is 32767
        ((Action)(() => s["A1"].SetString(tooLong))).Should().Throw<ResourceLimitExceededException>();
    }
}
