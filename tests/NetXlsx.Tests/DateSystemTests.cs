// Coverage for WorkbookOptions.DateSystem (design #15). Verifies the
// 1904 epoch is actually wired through to the workbook flag, the date
// serial, round-trip read-back, and that opening a file never clobbers
// the file's own epoch.

using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class DateSystemTests
{
    private static readonly DateTime SampleDate = new(2000, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

    // The 1900 and 1904 epochs differ by exactly 1462 days (4 years
    // including the spurious 1900 leap day).
    private const double EpochOffsetDays = 1462;

    [Fact]
    public void Default_Is_1900()
    {
        using var wb = Workbook.Create();
        wb.Underlying.IsDate1904().Should().BeFalse();
    }

    [Fact]
    public void Excel1900_Option_Leaves_Workbook_In_1900()
    {
        using var wb = Workbook.Create(new WorkbookOptions { DateSystem = DateSystem.Excel1900 });
        wb.Underlying.IsDate1904().Should().BeFalse();
    }

    [Fact]
    public void Excel1904_Option_Sets_Workbook_Flag()
    {
        using var wb = Workbook.Create(new WorkbookOptions { DateSystem = DateSystem.Excel1904 });
        wb.Underlying.IsDate1904().Should().BeTrue();
    }

    [Fact]
    public void Excel1904_Shifts_The_Serial_By_The_Epoch_Offset()
    {
        double serial1900, serial1904;

        using (var wb = Workbook.Create(new WorkbookOptions { DateSystem = DateSystem.Excel1900 }))
        {
            var c = wb.AddSheet("S")[1, 1];
            c.SetDate(SampleDate);
            serial1900 = c.Underlying.NumericCellValue;
        }
        using (var wb = Workbook.Create(new WorkbookOptions { DateSystem = DateSystem.Excel1904 }))
        {
            var c = wb.AddSheet("S")[1, 1];
            c.SetDate(SampleDate);
            serial1904 = c.Underlying.NumericCellValue;
        }

        (serial1900 - serial1904).Should().Be(EpochOffsetDays);
    }

    [Fact]
    public void Excel1904_Round_Trips_Flag_And_Value_Through_File()
    {
        var path = Path.Combine(Path.GetTempPath(), $"date1904-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create(new WorkbookOptions { DateSystem = DateSystem.Excel1904 }))
            {
                wb.AddSheet("S")[1, 1].SetDate(SampleDate);
                wb.Save(path);
            }
            // Reopen with default (1900) options: the file's own epoch must win.
            using (var wb = Workbook.Open(path))
            {
                wb.Underlying.IsDate1904().Should().BeTrue();
                wb["S"][1, 1].GetDate().Should().Be(SampleDate);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Open_Does_Not_Clobber_An_Existing_1904_File()
    {
        var path = Path.Combine(Path.GetTempPath(), $"date1904-open-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create(new WorkbookOptions { DateSystem = DateSystem.Excel1904 }))
            {
                wb.AddSheet("S")[1, 1].SetDate(SampleDate);
                wb.Save(path);
            }
            // Open explicitly requesting 1900 — the file is authoritative,
            // so the flag stays 1904 and the date still reads back correctly.
            using (var wb = Workbook.Open(path, new WorkbookOptions { DateSystem = DateSystem.Excel1900 }))
            {
                wb.Underlying.IsDate1904().Should().BeTrue();
                wb["S"][1, 1].GetDate().Should().Be(SampleDate);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Streaming_Honors_Excel1904()
    {
        var path = Path.Combine(Path.GetTempPath(), $"date1904-stream-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.CreateStreaming(new StreamingOptions { DateSystem = DateSystem.Excel1904 }))
            {
                var sheet = wb.AddSheet("S");
                sheet.AppendRow().Cell(1).SetDate(SampleDate);
                wb.Underlying.IsDate1904().Should().BeTrue();
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb.Underlying.IsDate1904().Should().BeTrue();
                wb["S"][1, 1].GetDate().Should().Be(SampleDate);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
