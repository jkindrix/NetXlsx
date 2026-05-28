using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles;

/// <summary>
/// Enforces the fixture-provenance rule (roadmap Process rules / decision I18):
/// every <c>.xlsx</c> under <c>Fixtures/</c> must have a sibling
/// <c>.fixture.md</c> file documenting its provenance. Failing this test
/// is the v1.0 ship-blocker pathway for ungoverned fixtures.
/// </summary>
public class FixtureProvenanceTests
{
    [Fact]
    public void Every_Xlsx_Fixture_Has_A_Sibling_Provenance_Doc()
    {
        var fixturesDir = Path.Combine(AppContext.BaseDirectory, "Fixtures");

        // The Fixtures/ directory is copied to the test output by the
        // csproj (it always contains at least README.md). If it's missing,
        // the fixture-copy infrastructure is broken — fail loud rather than
        // silently passing, which would let the provenance gate rot.
        Directory.Exists(fixturesDir).Should().BeTrue(
            "Fixtures/ must be copied to the test output dir (see NetXlsx.GoldenFiles.csproj). " +
            "A missing directory means the governance gate is no longer running.");

        var xlsxFiles = Directory.EnumerateFiles(fixturesDir, "*.xlsx", SearchOption.AllDirectories);
        var missing = xlsxFiles
            .Where(p => !File.Exists(Path.ChangeExtension(p, ".fixture.md")))
            .ToList();

        missing.Should().BeEmpty(
            "every fixture .xlsx requires a sibling .fixture.md per decision I18");
    }
}
