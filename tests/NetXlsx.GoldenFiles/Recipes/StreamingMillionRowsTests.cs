// Golden-file test for cookbook recipe 9 (StreamingMillionRows).
//
// Runs at a CI-friendly row count (5,000) — enough to exercise the
// streaming flush behavior past the 1,000-row window without dragging
// every test run by minutes. The recipe itself defaults to 250k for
// the runnable cookbook; a true 1M-row perf check belongs in
// benchmarks/, not the test suite.

using System;
using System.IO;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class StreamingMillionRowsTests
{
    private const int CiRowCount = 5_000;

    [Fact]
    public async Task Recipe_Streams_Many_Rows_And_File_Round_Trips()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-stream-{Guid.NewGuid():N}.xlsx");
        try
        {
            await StreamingMillionRows.Run(path, CiRowCount);

            // File exists and is non-trivial; SXSSF compression should
            // keep this well under a few MB for 5k rows × 20 cols.
            File.Exists(path).Should().BeTrue();
            new FileInfo(path).Length.Should().BeGreaterThan(10_000);

            using var wb = Workbook.Open(path);
            var sheet = wb[StreamingMillionRows.SheetName];

            // Header.
            sheet["A1"].GetString().Should().Be("Id");
            sheet["B1"].GetString().Should().Be("Col2");
            sheet["T1"].GetString().Should().Be("Col20");

            // Spot-check first, middle, and last data rows.
            sheet["A2"].GetNumber().Should().Be(1.0);
            sheet[CiRowCount / 2 + 1, 1].GetNumber().Should().Be(CiRowCount / 2);
            sheet[CiRowCount + 1, 1].GetNumber().Should().Be(CiRowCount);

            // Spot-check the mixed-type pattern from the recipe: at
            // column 3 (c % 3 == 0) we wrote (r*c)/7.0 as a double.
            // For r=1, c=3: 3/7.0 ≈ 0.428571…
            sheet["C2"].GetNumber().Should().BeApproximately(3.0 / 7.0, 1e-9);
            // At column 4 (c % 3 == 1) we wrote a string.
            sheet["D2"].GetString().Should().Be("r1-c4");
            // At column 5 (c % 3 == 2) we wrote r*c as an int.
            sheet["E2"].GetNumber().Should().Be(5.0);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
