using FluentAssertions;
using Xunit;

namespace NetXlsx.Tests;

/// <summary>
/// Scaffold smoke tests — confirm the test pipeline runs end-to-end before
/// any real implementation lands. Removed once real tests exist.
/// </summary>
public class ScaffoldSmokeTests
{
    [Fact]
    public void Xunit_Runs()
    {
        true.Should().BeTrue();
    }

    [Fact]
    public void FluentAssertions_Loaded()
    {
        var two = 1 + 1;
        two.Should().Be(2);
    }
}
