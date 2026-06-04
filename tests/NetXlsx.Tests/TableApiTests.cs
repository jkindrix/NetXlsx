// Coverage for the v1.1 Excel Tables (ListObject) slice: ISheet.AddTable,
// Tables, TryGetTable; ITable properties; file round-trip; name + range
// validation; style application.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class TableApiTests
{
    private static readonly string[] s_threeCols = new[] { "Region", "Revenue", "Margin" };
    private static readonly string[] s_twoCols = new[] { "Region", "Revenue" };

    // ---- Happy path ----------------------------------------------------

    [Fact]
    public void AddTable_Materializes_Columns_From_Header_Row()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("Region");
        sh["B1"].SetString("Revenue");
        sh["C1"].SetString("Margin");
        sh["A2"].SetString("EU");
        sh["B2"].SetNumber(100);
        sh["C2"].SetNumber(0.2);

        var t = sh.AddTable("A1:C2", "Sales");

        t.Name.Should().Be("Sales");
        t.DisplayName.Should().Be("Sales");
        t.Address.Should().Be("A1:C2");
        t.ColumnNames.Should().BeEquivalentTo(s_threeCols,
            opt => opt.WithStrictOrdering());
        t.HasTotalsRow.Should().BeFalse();
        t.StyleName.Should().BeNull();
        t.Sheet.Should().BeSameAs(sh);
    }

    [Fact]
    public void AddTable_Applies_Style_When_Provided()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H");
        sh["A2"].SetString("V");
        var t = sh.AddTable("A1:A2", "T", TableStyles.Medium2);
        t.StyleName.Should().Be("TableStyleMedium2");
    }

    [Fact]
    public void Tables_Returns_Empty_When_No_Tables_Defined()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.Tables.Should().BeEmpty();
    }

    [Fact]
    public void Tables_Snapshot_Contains_Added_Tables()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H1"); sh["A2"].SetString("v1");
        sh["B1"].SetString("H2"); sh["B2"].SetString("v2");
        sh.AddTable("A1:A2", "First");
        sh.AddTable("B1:B2", "Second");
        sh.Tables.Should().HaveCount(2);
        sh.Tables.Should().Contain(t => t.Name == "First")
                 .And.Contain(t => t.Name == "Second");
    }

    [Fact]
    public void TryGetTable_Finds_By_Name_Case_Insensitive()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H"); sh["A2"].SetString("v");
        sh.AddTable("A1:A2", "MyTable");

        sh.TryGetTable("mytable", out var t).Should().BeTrue();
        t!.Name.Should().Be("MyTable");

        sh.TryGetTable("Missing", out _).Should().BeFalse();
    }

    [Fact]
    public void DisplayName_And_StyleName_Are_Settable()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H"); sh["A2"].SetString("v");
        var t = sh.AddTable("A1:A2", "T");

        t.DisplayName = "Friendly";
        t.DisplayName.Should().Be("Friendly");

        t.StyleName = TableStyles.Dark9;
        t.StyleName.Should().Be("TableStyleDark9");

        t.StyleName = null;
        t.StyleName.Should().BeNull();
    }

    // ---- Validation ----------------------------------------------------

    [Fact]
    public void AddTable_Rejects_Null_Or_Empty_Name()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H"); sh["A2"].SetString("v");
        Action a = () => sh.AddTable("A1:A2", null!);
        a.Should().Throw<ArgumentNullException>();
        Action b = () => sh.AddTable("A1:A2", "");
        b.Should().Throw<ArgumentException>();
    }

    [Theory]
    [InlineData("1Table", "must start with a letter or underscore")]
    [InlineData("Bad Name", "invalid character ' '")]
    [InlineData("A1", "collides with an A1-style cell address")]
    [InlineData("AB123", "collides with an A1-style cell address")]
    public void AddTable_Rejects_Invalid_Names(string name, string messageSubstring)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H"); sh["A2"].SetString("v");
        Action act = () => sh.AddTable("A1:A2", name);
        act.Should().Throw<ArgumentException>().WithMessage($"*{messageSubstring}*");
    }

    [Fact]
    public void AddTable_Rejects_Duplicate_Name_Workbook_Wide()
    {
        using var wb = Workbook.Create();
        var s1 = wb.AddSheet("S1");
        var s2 = wb.AddSheet("S2");
        s1["A1"].SetString("H"); s1["A2"].SetString("v");
        s2["A1"].SetString("H"); s2["A2"].SetString("v");
        s1.AddTable("A1:A2", "Shared");

        Action act = () => s2.AddTable("A1:A2", "Shared");
        act.Should().Throw<ArgumentException>().WithMessage("*already exists*");
    }

    [Fact]
    public void AddTable_Rejects_Collision_With_Named_Range()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H"); sh["A2"].SetString("v");
        wb.AddNamedRange("Taken", "S!$A$1");

        Action act = () => sh.AddTable("A1:A2", "Taken");
        act.Should().Throw<ArgumentException>().WithMessage("*named range*");
    }

    [Fact]
    public void AddTable_Rejects_Single_Row_Range()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H");
        Action act = () => sh.AddTable("A1:C1", "T");
        act.Should().Throw<ArgumentException>().WithMessage("*at least one data row*");
    }

    [Fact]
    public void AddTable_Rejects_Empty_Header_Cells()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        // A1 has header, B1 does not
        sh["A1"].SetString("H");
        sh["A2"].SetString("v"); sh["B2"].SetString("v");
        Action act = () => sh.AddTable("A1:B2", "T");
        act.Should().Throw<ArgumentException>().WithMessage("*empty or non-string*");
    }

    [Fact]
    public void AddTable_Rejects_Non_String_Header_Cells()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetNumber(1.0);    // numeric header — not allowed
        sh["A2"].SetString("v");
        Action act = () => sh.AddTable("A1:A2", "T");
        act.Should().Throw<ArgumentException>().WithMessage("*empty or non-string*");
    }

    [Fact]
    public void AddTable_Rejects_Duplicate_Headers()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("Same");
        sh["B1"].SetString("same");   // case-insensitive duplicate
        sh["A2"].SetString("v"); sh["B2"].SetString("v");
        Action act = () => sh.AddTable("A1:B2", "T");
        act.Should().Throw<ArgumentException>().WithMessage("*Duplicate header*");
    }

    // ---- RemoveTable (v1.2 / I-63) ------------------------------------

    [Fact]
    public void RemoveTable_Drops_The_Table_From_Tables_Snapshot()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H"); sh["A2"].SetString("v");
        var t = sh.AddTable("A1:A2", "T");

        sh.Tables.Should().HaveCount(1);
        sh.RemoveTable(t);
        sh.Tables.Should().BeEmpty();
    }

    [Fact]
    public void RemoveTable_Allows_Subsequent_AddTable_With_Same_Name()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H"); sh["A2"].SetString("v");

        var t = sh.AddTable("A1:A2", "Reusable");
        sh.RemoveTable(t);

        // The name is freed; we can register a fresh table with the
        // same codename without a uniqueness collision.
        Action act = () => sh.AddTable("A1:A2", "Reusable");
        act.Should().NotThrow();
        sh.Tables.Should().HaveCount(1);
        sh.Tables[0].Name.Should().Be("Reusable");
    }

    [Fact]
    public void RemoveTable_Rejects_Null()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.RemoveTable(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void RemoveTable_Rejects_Table_From_Different_Sheet()
    {
        using var wb = Workbook.Create();
        var s1 = wb.AddSheet("S1");
        var s2 = wb.AddSheet("S2");
        s1["A1"].SetString("H"); s1["A2"].SetString("v");
        s2["A1"].SetString("H"); s2["A2"].SetString("v");

        var t1 = s1.AddTable("A1:A2", "SheetOneTable");
        Action act = () => s2.RemoveTable(t1);
        act.Should().Throw<ArgumentException>().WithMessage("*does not belong*");
    }

    [Fact]
    public void RemoveTable_Idempotent_When_Called_Twice_Throws_Second_Time()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H"); sh["A2"].SetString("v");

        var t = sh.AddTable("A1:A2", "T");
        sh.RemoveTable(t);

        // The handle is now stale — its underlying part no longer has a
        // relation to the sheet. A second remove must fail loudly, not
        // silently no-op (silent-no-op would mask caller bugs).
        Action act = () => sh.RemoveTable(t);
        act.Should().Throw<ArgumentException>().WithMessage("*does not belong*");
    }

    [Fact]
    public void RemoveTable_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"rmtbl-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh.AppendRow().Set(1, "H1").Set(2, "H2");
                sh.AppendRow().Set(1, "a").Set(2, "b");
                var kept = sh.AddTable("A1:A2", "Kept");
                var dropped = sh.AddTable("B1:B2", "Dropped");
                sh.RemoveTable(dropped);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var sh = wb["S"];
                sh.Tables.Should().HaveCount(1);
                sh.Tables[0].Name.Should().Be("Kept");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Totals row (v1.2 / I-64) -------------------------------------

    [Fact]
    public void AddTotalsRow_Extends_Range_And_Flips_HasTotalsRow()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("Region"); sh["B1"].SetString("Revenue");
        sh["A2"].SetString("EU"); sh["B2"].SetNumber(100);
        sh["A3"].SetString("US"); sh["B3"].SetNumber(200);
        var t = sh.AddTable("A1:B3", "Sales");

        t.HasTotalsRow.Should().BeFalse();
        t.Address.Should().Be("A1:B3");

        t.AddTotalsRow();

        t.HasTotalsRow.Should().BeTrue();
        t.Address.Should().Be("A1:B4");
    }

    [Fact]
    public void AddTotalsRow_Is_Idempotent()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H"); sh["A2"].SetString("v");
        var t = sh.AddTable("A1:A2", "T");
        t.AddTotalsRow();
        t.AddTotalsRow();
        t.HasTotalsRow.Should().BeTrue();
        t.Address.Should().Be("A1:A3");   // not "A1:A4"
    }

    [Fact]
    public void RemoveTotalsRow_Shrinks_Range_And_Flips_HasTotalsRow()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H"); sh["A2"].SetString("v");
        var t = sh.AddTable("A1:A2", "T");
        t.AddTotalsRow();
        t.SetColumnTotal("H", TotalsRowFunction.Count);

        t.RemoveTotalsRow();

        t.HasTotalsRow.Should().BeFalse();
        t.Address.Should().Be("A1:A2");
        // Persisted table metadata cleared
        TotalsRowFunctionOf(wb, 0).Should().Be("none");
    }

    [Fact]
    public void RemoveTotalsRow_Is_Safe_On_Table_Without_Totals()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("H"); sh["A2"].SetString("v");
        var t = sh.AddTable("A1:A2", "T");
        Action act = () => t.RemoveTotalsRow();
        act.Should().NotThrow();
        t.HasTotalsRow.Should().BeFalse();
    }

    [Fact]
    public void SetColumnTotal_Writes_SUBTOTAL_Formula_And_Function_Metadata()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("Region"); sh["B1"].SetString("Revenue");
        sh["A2"].SetString("EU"); sh["B2"].SetNumber(100);
        sh["A3"].SetString("US"); sh["B3"].SetNumber(200);
        var t = sh.AddTable("A1:B3", "Sales");
        t.AddTotalsRow();

        t.SetColumnTotal("Revenue", TotalsRowFunction.Sum);

        // Cell B4 carries the SUBTOTAL formula.
        sh["B4"].GetFormula().Should().Be("=SUBTOTAL(109,Sales[Revenue])");
        // Persisted table metadata set to sum.
        TotalsRowFunctionOf(wb, 1).Should().Be("sum");
    }

    [Theory]
    [InlineData(TotalsRowFunction.Sum, 109)]
    [InlineData(TotalsRowFunction.Average, 101)]
    [InlineData(TotalsRowFunction.Min, 105)]
    [InlineData(TotalsRowFunction.Max, 104)]
    [InlineData(TotalsRowFunction.Count, 103)]
    [InlineData(TotalsRowFunction.CountNumbers, 102)]
    [InlineData(TotalsRowFunction.StdDev, 107)]
    [InlineData(TotalsRowFunction.Var, 110)]
    public void SetColumnTotal_Each_Builtin_Function_Emits_Correct_SUBTOTAL_Code(TotalsRowFunction f, int code)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("V"); sh["A2"].SetNumber(1); sh["A3"].SetNumber(2);
        var t = sh.AddTable("A1:A3", "Nums");
        t.AddTotalsRow();
        t.SetColumnTotal("V", f);
        sh["A4"].GetFormula().Should().Be($"=SUBTOTAL({code},Nums[V])");
    }

    [Fact]
    public void SetColumnTotal_Custom_Writes_Supplied_Formula()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("V"); sh["A2"].SetNumber(1); sh["A3"].SetNumber(2);
        var t = sh.AddTable("A1:A3", "Nums");
        t.AddTotalsRow();
        t.SetColumnTotal("V", "SUMPRODUCT(Nums[V])");
        sh["A4"].GetFormula().Should().Be("=SUMPRODUCT(Nums[V])");
    }

    [Fact]
    public void SetColumnTotal_Custom_Strips_Leading_Equals()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("V"); sh["A2"].SetNumber(1); sh["A3"].SetNumber(2);
        var t = sh.AddTable("A1:A3", "Nums");
        t.AddTotalsRow();
        t.SetColumnTotal("V", "=SUM(Nums[V])*2");
        sh["A4"].GetFormula().Should().Be("=SUM(Nums[V])*2");
    }

    [Fact]
    public void SetColumnTotalLabel_Writes_Plain_Text()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("Region"); sh["B1"].SetString("Revenue");
        sh["A2"].SetString("EU"); sh["B2"].SetNumber(100);
        var t = sh.AddTable("A1:B2", "Sales");
        t.AddTotalsRow();

        t.SetColumnTotalLabel("Region", "Total");

        sh["A3"].GetString().Should().Be("Total");
        // Label sets function back to none (label takes precedence over function).
        TotalsRowFunctionOf(wb, 0).Should().Be("none");
    }

    [Fact]
    public void SetColumnTotal_Requires_HasTotalsRow()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("V"); sh["A2"].SetNumber(1);
        var t = sh.AddTable("A1:A2", "T");

        Action act = () => t.SetColumnTotal("V", TotalsRowFunction.Sum);
        act.Should().Throw<InvalidOperationException>().WithMessage("*AddTotalsRow*");
    }

    [Fact]
    public void SetColumnTotal_Rejects_Unknown_Column()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("V"); sh["A2"].SetNumber(1);
        var t = sh.AddTable("A1:A2", "T");
        t.AddTotalsRow();

        Action act = () => t.SetColumnTotal("DoesNotExist", TotalsRowFunction.Sum);
        act.Should().Throw<ArgumentException>().WithMessage("*not part of*");
    }

    [Fact]
    public void SetColumnTotal_Rejects_Custom_Via_Enum_Overload()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("V"); sh["A2"].SetNumber(1);
        var t = sh.AddTable("A1:A2", "T");
        t.AddTotalsRow();

        Action act = () => t.SetColumnTotal("V", TotalsRowFunction.Custom);
        act.Should().Throw<ArgumentException>().WithMessage("*customFormula*");
    }

    [Fact]
    public void TotalsRow_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"totals-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh["A1"].SetString("Region"); sh["B1"].SetString("Revenue");
                sh["A2"].SetString("EU"); sh["B2"].SetNumber(100);
                sh["A3"].SetString("US"); sh["B3"].SetNumber(200);
                var t = sh.AddTable("A1:B3", "Sales");
                t.AddTotalsRow();
                t.SetColumnTotalLabel("Region", "Total");
                t.SetColumnTotal("Revenue", TotalsRowFunction.Sum);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var sh = wb["S"];
                var t = sh.Tables[0];
                t.HasTotalsRow.Should().BeTrue();
                t.Address.Should().Be("A1:B4");
                sh["A4"].GetString().Should().Be("Total");
                sh["B4"].GetFormula().Should().Be("=SUBTOTAL(109,Sales[Revenue])");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- File round-trip ----------------------------------------------

    [Fact]
    public void AddTable_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"table-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh["A1"].SetString("Region");
                sh["B1"].SetString("Revenue");
                sh["A2"].SetString("EU"); sh["B2"].SetNumber(100);
                sh["A3"].SetString("US"); sh["B3"].SetNumber(200);
                sh.AddTable("A1:B3", "Sales", TableStyles.Medium2);
                wb.Save(path);
            }

            using (var wb = Workbook.Open(path))
            {
                var sh = wb["S"];
                sh.Tables.Should().HaveCount(1);
                var t = sh.Tables[0];
                t.Name.Should().Be("Sales");
                t.DisplayName.Should().Be("Sales");
                t.Address.Should().Be("A1:B3");
                t.ColumnNames.Should().BeEquivalentTo(s_twoCols,
                    opt => opt.WithStrictOrdering());
                t.StyleName.Should().Be("TableStyleMedium2");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- helpers ------------------------------------------------------

    /// <summary>
    /// tableColumn/@totalsRowFunction of the first table part; the OOXML
    /// default (absent attribute) is "none".
    /// </summary>
    private static string TotalsRowFunctionOf(IWorkbook wb, int columnIndex)
        => (string?)SavedOoxml.Part(wb, "xl/tables/table1.xml").Root!
            .Element(SavedOoxml.Main + "tableColumns")!
            .Elements(SavedOoxml.Main + "tableColumn")
            .ElementAt(columnIndex)
            .Attribute("totalsRowFunction") ?? "none";

}
