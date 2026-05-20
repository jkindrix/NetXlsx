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
        if (!Directory.Exists(fixturesDir))
        {
            // Scaffold: no fixtures yet. Test passes trivially until fixtures
            // land. Once any .xlsx exists, the assertion below applies.
            return;
        }

        var xlsxFiles = Directory.EnumerateFiles(fixturesDir, "*.xlsx", SearchOption.AllDirectories);
        var missing = xlsxFiles
            .Where(p => !File.Exists(Path.ChangeExtension(p, ".fixture.md")))
            .ToList();

        missing.Should().BeEmpty(
            "every fixture .xlsx requires a sibling .fixture.md per decision I18");
    }
}
