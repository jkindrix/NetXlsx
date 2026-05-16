// Workbook.SuggestSheetName — design line 160.

using System;
using FluentAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class SuggestSheetNameTests
{
    [Fact]
    public void Returns_Proposed_Verbatim_When_No_Collision()
    {
        using var wb = Workbook.Create();
        Workbook.SuggestSheetName(wb, "Sales").Should().Be("Sales");
    }

    [Fact]
    public void Appends_Numeric_Suffix_On_Collision()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Sales");
        Workbook.SuggestSheetName(wb, "Sales").Should().Be("Sales (2)");
    }

    [Fact]
    public void Walks_Suffixes_Until_Unused()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Sales");
        wb.AddSheet("Sales (2)");
        wb.AddSheet("Sales (3)");
        Workbook.SuggestSheetName(wb, "Sales").Should().Be("Sales (4)");
    }

    [Fact]
    public void Is_Case_Insensitive_For_Collision_Check()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Sales");
        // "SALES" collides with "Sales" (Excel sheet-name lookup is
        // case-insensitive — matches AddSheet's duplicate detection).
        Workbook.SuggestSheetName(wb, "SALES").Should().Be("SALES (2)");
    }

    [Fact]
    public void Truncates_To_31_Chars_Preserving_Suffix()
    {
        using var wb = Workbook.Create();
        var thirtyOne = new string('x', 31);
        wb.AddSheet(thirtyOne);

        var result = Workbook.SuggestSheetName(wb, thirtyOne);
        result.Length.Should().BeLessOrEqualTo(31);
        result.Should().EndWith(" (2)");
    }

    [Fact]
    public void Sanitizes_Invalid_Characters()
    {
        using var wb = Workbook.Create();
        Workbook.SuggestSheetName(wb, "Bad/Name").Should().Be("Bad_Name");
    }

    [Fact]
    public void Throws_On_Null_Inputs()
    {
        using var wb = Workbook.Create();
        Action a = () => Workbook.SuggestSheetName(null!, "x");
        Action b = () => Workbook.SuggestSheetName(wb, null!);
        a.Should().Throw<ArgumentNullException>();
        b.Should().Throw<ArgumentNullException>();
    }
}
