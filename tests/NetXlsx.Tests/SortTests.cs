using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class SortTests
{
    [Fact]
    public void SortRange_Ascending_String_Column()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("Charlie");
        s["A2"].SetString("Alice");
        s["A3"].SetString("Bob");

        s.SortRange("A1:A3", SortKey.Asc(1));

        s["A1"].GetString().Should().Be("Alice");
        s["A2"].GetString().Should().Be("Bob");
        s["A3"].GetString().Should().Be("Charlie");
    }

    [Fact]
    public void SortRange_Descending_Numeric_Column()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetNumber(10);
        s["A2"].SetNumber(30);
        s["A3"].SetNumber(20);

        s.SortRange("A1:A3", SortKey.Desc(1));

        s["A1"].GetNumber().Should().Be(30);
        s["A2"].GetNumber().Should().Be(20);
        s["A3"].GetNumber().Should().Be(10);
    }

    [Fact]
    public void SortRange_Multi_Column_Key()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("B"); s["B1"].SetNumber(2);
        s["A2"].SetString("A"); s["B2"].SetNumber(3);
        s["A3"].SetString("A"); s["B3"].SetNumber(1);

        s.SortRange("A1:B3", SortKey.Asc(1), SortKey.Asc(2));

        s["A1"].GetString().Should().Be("A");
        s["B1"].GetNumber().Should().Be(1);
        s["A2"].GetString().Should().Be("A");
        s["B2"].GetNumber().Should().Be(3);
        s["A3"].GetString().Should().Be("B");
        s["B3"].GetNumber().Should().Be(2);
    }

    [Fact]
    public void SortRange_Preserves_Styles()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var bold = new CellStyle { Bold = true };
        s["A1"].SetString("B"); s["A1"].Style(bold);
        s["A2"].SetString("A");

        s.SortRange("A1:A2", SortKey.Asc(1));

        s["A1"].GetString().Should().Be("A");
        s["A2"].GetString().Should().Be("B");
        // The bold style moved with "B"
        s["A2"].GetStyle().Bold.Should().BeTrue();
    }

    [Fact]
    public void SortRange_Blanks_Sort_Last_Ascending()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A2"].SetNumber(1);
        s["A3"].SetNumber(3);
        // A1 is blank

        s.SortRange("A1:A3", SortKey.Asc(1));

        s["A1"].GetNumber().Should().Be(1);
        s["A2"].GetNumber().Should().Be(3);
    }

    [Fact]
    public void SortRange_Single_Row_Is_Noop()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("only");

        s.SortRange("A1:A1", SortKey.Asc(1));

        s["A1"].GetString().Should().Be("only");
    }

    [Fact]
    public void SortRange_Rejects_Null_Range()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.SortRange(null!, SortKey.Asc(1));
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void SortRange_Rejects_Empty_Keys()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.SortRange("A1:A3");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void SortKey_Asc_Creates_Ascending()
    {
        var key = SortKey.Asc(3);
        key.Column.Should().Be(3);
        key.Ascending.Should().BeTrue();
    }

    [Fact]
    public void SortKey_Desc_Creates_Descending()
    {
        var key = SortKey.Desc(2);
        key.Column.Should().Be(2);
        key.Ascending.Should().BeFalse();
    }

    [Fact]
    public void SortKey_Rejects_Zero_Column()
    {
        Action act = () => SortKey.Asc(0);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void SortRange_Survives_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("S");
            s["A1"].SetNumber(30);
            s["A2"].SetNumber(10);
            s["A3"].SetNumber(20);
            s.SortRange("A1:A3", SortKey.Asc(1));
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        opened["S"]["A1"].GetNumber().Should().Be(10);
        opened["S"]["A2"].GetNumber().Should().Be(20);
        opened["S"]["A3"].GetNumber().Should().Be(30);
    }
}
