// Coverage for the A1 parser per design §6.10.

using System;
using FluentAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class CellAddressTests
{
    [Theory]
    [InlineData("A1", 1, 1)]
    [InlineData("B2", 2, 2)]
    [InlineData("Z1", 1, 26)]
    [InlineData("AA1", 1, 27)]
    [InlineData("AZ1", 1, 52)]
    [InlineData("BA1", 1, 53)]
    [InlineData("ZZ1", 1, 702)]
    [InlineData("AAA1", 1, 703)]
    [InlineData("XFD1", 1, 16384)]
    [InlineData("A1048576", 1048576, 1)]
    public void Parse_Accepts_Canonical_A1(string input, int expectedRow, int expectedCol)
    {
        var (row, col) = CellAddress.Parse(input);
        row.Should().Be(expectedRow);
        col.Should().Be(expectedCol);
    }

    [Theory]
    [InlineData("a1", 1, 1)]
    [InlineData("aa10", 10, 27)]
    public void Parse_Is_Case_Insensitive(string input, int expectedRow, int expectedCol)
    {
        var (row, col) = CellAddress.Parse(input);
        row.Should().Be(expectedRow);
        col.Should().Be(expectedCol);
    }

    [Theory]
    [InlineData("$A$1", 1, 1)]
    [InlineData("$A1", 1, 1)]
    [InlineData("A$1", 1, 1)]
    public void Parse_Strips_Dollar_Signs(string input, int expectedRow, int expectedCol)
    {
        var (row, col) = CellAddress.Parse(input);
        row.Should().Be(expectedRow);
        col.Should().Be(expectedCol);
    }

    [Theory]
    [InlineData("Sheet1!A1", "sheet-qualified")]
    [InlineData("A1:C10", "range")]
    [InlineData("A:A", "range")]
    [InlineData("1:1", "range")]
    [InlineData("", "empty")]
    [InlineData("1", "missing column")]
    [InlineData("A", "missing row")]
    [InlineData("A0", "row index must be 1")]
    [InlineData("AAAAA1", "exceeds Excel maximum")]
    [InlineData("A1048577", "exceeds Excel maximum")]
    [InlineData("A1B", "trailing characters")]
    [InlineData("A 1", "missing row digits")]   // space terminates the column scan; nothing for row
    [InlineData("123A", "missing column letters")]   // digit-led; no column letters consumed
    public void Parse_Rejects_Invalid_Forms(string input, string reasonFragment)
    {
        var act = () => CellAddress.Parse(input);
        act.Should().Throw<InvalidCellAddressException>()
            .Where(ex => ex.Message.Contains(reasonFragment, System.StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(1, 1, "A1")]
    [InlineData(10, 27, "AA10")]
    [InlineData(1, 16384, "XFD1")]
    [InlineData(1048576, 1, "A1048576")]
    public void Format_Roundtrips_Through_Parse(int row, int col, string expected)
    {
        CellAddress.Format(row, col).Should().Be(expected);
        var (r, c) = CellAddress.Parse(expected);
        r.Should().Be(row);
        c.Should().Be(col);
    }

    [Fact]
    public void TryParse_Returns_False_Without_Throwing()
    {
        CellAddress.TryParse("not a cell", out int row, out int col).Should().BeFalse();
        row.Should().Be(0);
        col.Should().Be(0);
    }

    // ---- ParseRange / FormatRange (v0.6 sub-slice A) ------------------

    [Theory]
    [InlineData("A1:C3", 1, 1, 3, 3)]
    [InlineData("A1:A1", 1, 1, 1, 1)]
    [InlineData("B2:D5", 2, 2, 5, 4)]
    [InlineData("a1:c3", 1, 1, 3, 3)]                 // case-insensitive
    [InlineData("$A$1:$C$3", 1, 1, 3, 3)]             // dollar signs stripped
    public void ParseRange_Accepts_Bounded_Forms(string input, int r1, int c1, int r2, int c2)
    {
        var (row1, col1, row2, col2) = CellAddress.ParseRange(input);
        row1.Should().Be(r1);
        col1.Should().Be(c1);
        row2.Should().Be(r2);
        col2.Should().Be(c2);
    }

    [Fact]
    public void ParseRange_Single_Cell_Form_Treated_As_1x1_Range()
    {
        var (r1, c1, r2, c2) = CellAddress.ParseRange("A1");
        r1.Should().Be(1); c1.Should().Be(1); r2.Should().Be(1); c2.Should().Be(1);
    }

    [Fact]
    public void ParseRange_Normalizes_Inverted_Corners()
    {
        // "C3:A1" -> top-left A1, bottom-right C3.
        var (r1, c1, r2, c2) = CellAddress.ParseRange("C3:A1");
        r1.Should().Be(1); c1.Should().Be(1); r2.Should().Be(3); c2.Should().Be(3);
    }

    [Theory]
    [InlineData("Sheet1!A1:B2")]    // sheet-qualified — not accepted by Range
    [InlineData("garbage")]
    [InlineData(":A1")]              // empty left side
    [InlineData("A1:")]              // empty right side
    [InlineData("A:5")]              // mixed column-then-row
    [InlineData("1:A")]              // mixed row-then-column
    public void ParseRange_Rejects_Invalid_Forms(string input)
    {
        ((Action)(() => CellAddress.ParseRange(input))).Should()
            .Throw<InvalidCellAddressException>();
    }

    [Theory]
    [InlineData("A:A",   1,            1, CellAddress.MaxRow,    1)]
    [InlineData("A:C",   1,            1, CellAddress.MaxRow,    3)]
    [InlineData("AA:AB", 1,           27, CellAddress.MaxRow,   28)]
    [InlineData("1:1",   1,            1, 1,                     CellAddress.MaxColumn)]
    [InlineData("1:5",   1,            1, 5,                     CellAddress.MaxColumn)]
    [InlineData("C:A",   1,            1, CellAddress.MaxRow,    3)]   // normalized
    [InlineData("$A:$A", 1,            1, CellAddress.MaxRow,    1)]   // dollar-sign tolerance
    [InlineData("$1:$3", 1,            1, 3,                     CellAddress.MaxColumn)]
    public void ParseRange_Expands_WholeRow_And_WholeColumn_Shorthand(string input,
        int r1, int c1, int r2, int c2)
    {
        var (row1, col1, row2, col2) = CellAddress.ParseRange(input);
        row1.Should().Be(r1);
        col1.Should().Be(c1);
        row2.Should().Be(r2);
        col2.Should().Be(c2);
    }

    [Theory]
    [InlineData(1, 1, 1, 1, "A1")]          // 1x1 -> single cell
    [InlineData(1, 1, 3, 3, "A1:C3")]
    [InlineData(2, 27, 10, 28, "AA2:AB10")]
    public void FormatRange_Produces_Canonical_Form(int r1, int c1, int r2, int c2, string expected)
    {
        CellAddress.FormatRange(r1, c1, r2, c2).Should().Be(expected);
    }
}
