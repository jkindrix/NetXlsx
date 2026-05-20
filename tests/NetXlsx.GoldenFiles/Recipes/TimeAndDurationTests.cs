// Golden-file test for cookbook recipe 12 (TimeAndDuration).

using System;
using System.IO;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class TimeAndDurationTests
{
    [Fact]
    public async Task Recipe_Roundtrips_All_DateTime_Kinds_With_Correct_Format_Hints()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-time-{Guid.NewGuid():N}.xlsx");
        try
        {
            await TimeAndDuration.Run(path);

            using var wb = Workbook.Open(path);
            var sheet = wb[TimeAndDuration.SheetName];

            // Header.
            sheet["A1"].GetString().Should().Be("Kind");

            // DateTime row.
            sheet["A2"].GetString().Should().Be("DateTime");
            sheet["B2"].Kind.Should().Be(CellKind.Date);
            sheet["B2"].GetDate().Should().Be(new DateTime(2026, 5, 16, 9, 30, 15, DateTimeKind.Unspecified));

            // DateOnly row.
            sheet["A3"].GetString().Should().Be("DateOnly");
            sheet["B3"].Kind.Should().Be(CellKind.Date);
            sheet["B3"].GetDateOnly().Should().Be(new DateOnly(2026, 5, 16));

            // TimeOnly row. NPOI's IsCellDateFormatted treats h:mm:ss as
            // date-shaped, so the kind is Date (per design §7.10's
            // definition of CellKind.Date — "numeric value styled with
            // a date number format" — which time formats satisfy).
            sheet["A4"].GetString().Should().Be("TimeOnly");
            sheet["B4"].Kind.Should().Be(CellKind.Date);
            sheet["B4"].GetTime().Should().Be(new TimeOnly(9, 30, 15));

            // Duration < 24h.
            sheet["A5"].GetString().Should().Be("Duration (4h 15m)");
            sheet["B5"].GetDuration()!.Value.Should().BeCloseTo(
                TimeSpan.FromMinutes(255), TimeSpan.FromMicroseconds(1));

            // Duration > 24h — the value, not just modulo-24h.
            sheet["A6"].GetString().Should().Be("Duration (26h)");
            sheet["B6"].GetDuration()!.Value.Should().BeCloseTo(
                TimeSpan.FromHours(26), TimeSpan.FromMicroseconds(1));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task Recipe_Cells_Carry_Number_Format_Style_For_Excel_Display()
    {
        // The default-style application is what makes Excel show the
        // cells AS dates / times instead of as raw doubles. Verify the
        // format string survives the round trip through NPOI.
        var path = Path.Combine(Path.GetTempPath(), $"golden-time-fmt-{Guid.NewGuid():N}.xlsx");
        try
        {
            await TimeAndDuration.Run(path);

            using var wb = Workbook.Open(path);
            var sheet = wb[TimeAndDuration.SheetName];

            // DateTime cell — yyyy-mm-dd hh:mm:ss.
            sheet["B2"].Underlying.CellStyle.GetDataFormatString().Should().Be("yyyy-mm-dd hh:mm:ss");
            // DateOnly cell — yyyy-mm-dd (I-19).
            sheet["B3"].Underlying.CellStyle.GetDataFormatString().Should().Be("yyyy-mm-dd");
            // TimeOnly cell — h:mm:ss.
            sheet["B4"].Underlying.CellStyle.GetDataFormatString().Should().Be("h:mm:ss");
            // Duration cells — elapsed-time format [h]:mm:ss (§7.9).
            sheet["B5"].Underlying.CellStyle.GetDataFormatString().Should().Be("[h]:mm:ss");
            sheet["B6"].Underlying.CellStyle.GetDataFormatString().Should().Be("[h]:mm:ss");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
