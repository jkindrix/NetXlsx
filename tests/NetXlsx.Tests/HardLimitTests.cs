// Excel hard-limit enforcement on write (decision #37 / §7.6).

using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class HardLimitTests
{
    [Fact]
    public void SetString_Accepts_Cell_Text_Up_To_32767_Chars()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var atLimit = new string('a', 32_767);

        sheet["A1"].SetString(atLimit);
        sheet["A1"].GetString().Should().Be(atLimit);
    }

    [Fact]
    public void SetString_Throws_ResourceLimitExceeded_Above_32767_Chars()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var overLimit = new string('a', 32_768);

        var act = () => sheet["A1"].SetString(overLimit);
        act.Should().Throw<ResourceLimitExceededException>()
            .Where(ex => ex.LimitName == "cell text length"
                      && ex.Limit == 32_767
                      && ex.Actual == 32_768);
    }
}
