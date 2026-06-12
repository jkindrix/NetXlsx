// R-9: sheet-name validation tightened to Excel's actual rule set. The
// pre-fix gaps — all probe-verified accepted on 2026-06-10: colon ("a:b",
// which collides with 3D-reference syntax), leading/trailing apostrophe
// ("'Leading" / "Trailing'", which LibreOffice silently renames to
// "Sheet1" on resave), and the reserved name "History". One shared char
// array feeds IsValidSheetName / SanitizeSheetName / ValidateSheetName,
// so the three surfaces are pinned together here.

using System;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class SheetNameValidationTests
{
    [Theory]
    [InlineData("a:b")]        // colon — the worst pre-fix gap (3D-reference collision)
    [InlineData("'Leading")]
    [InlineData("Trailing'")]
    [InlineData("History")]
    [InlineData("history")]    // reservation is case-insensitive
    [InlineData("HISTORY")]
    public void AddSheet_Rejects_The_R9_Gap_Set(string name)
    {
        using var wb = Workbook.Create();
        Action act = () => wb.AddSheet(name);
        act.Should().Throw<SheetNameException>();
    }

    [Theory]
    [InlineData("a:b", false)]
    [InlineData("'Leading", false)]
    [InlineData("Trailing'", false)]
    [InlineData("History", false)]
    [InlineData("history", false)]
    [InlineData("O'Brien", true)]      // interior apostrophe is legal
    [InlineData("History 2026", true)] // reservation is exact-name only
    [InlineData("Sales", true)]
    public void IsValidSheetName_Matches_Excels_Rule_Set(string name, bool valid)
    {
        Workbook.IsValidSheetName(name).Should().Be(valid);
    }

    [Theory]
    [InlineData("a:b", "a_b")]
    [InlineData("'Leading", "_Leading")]
    [InlineData("Trailing'", "Trailing_")]
    [InlineData("History", "History_")]
    [InlineData("history", "history_")]
    [InlineData("''", "__")]
    [InlineData("O'Brien", "O'Brien")]
    public void SanitizeSheetName_Produces_A_Valid_Name(string input, string expected)
    {
        var sanitized = Workbook.SanitizeSheetName(input);
        sanitized.Should().Be(expected);
        Workbook.IsValidSheetName(sanitized).Should().BeTrue();
    }

    [Fact]
    public void SanitizeSheetName_Fixes_Apostrophe_Exposed_By_Truncation()
    {
        // 31st char is an apostrophe followed by more text — truncation
        // exposes it as the new trailing char; the fix must run after.
        var input = new string('a', 30) + "'suffix";
        var sanitized = Workbook.SanitizeSheetName(input);
        sanitized.Should().HaveLength(31);
        sanitized.Should().EndWith("_");
        Workbook.IsValidSheetName(sanitized).Should().BeTrue();
    }

    [Fact]
    public void Valid_Names_Still_Pass_End_To_End()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("O'Brien");
        wb.AddSheet("History 2026");
        wb["O'Brien"].Name.Should().Be("O'Brien");
    }
}
