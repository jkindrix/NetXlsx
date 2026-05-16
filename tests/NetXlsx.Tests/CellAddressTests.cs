// Coverage for the A1 parser per design §6.10.

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
}
