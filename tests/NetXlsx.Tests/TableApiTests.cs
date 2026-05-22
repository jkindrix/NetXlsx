// Coverage for the v1.1 Excel Tables (ListObject) slice: ISheet.AddTable,
// Tables, TryGetTable; ITable properties; file round-trip; name + range
// validation; style application.

using System;
using System.IO;
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
}
