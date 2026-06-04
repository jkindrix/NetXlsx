// Coverage for WorkbookOptions.DateSystem (design #15). Verifies the
// 1904 epoch is actually wired through to the workbook flag, the date
// serial, round-trip read-back, and that opening a file never clobbers
// the file's own epoch.
//
// The epoch flag and raw serial have no public read-back, so tests
// assert on the persisted workbookPr/@date1904 and the cell <v> via
// SavedOoxml — engine-agnostic, no .Underlying reach-through (I-82
// cutover phase 1).

using System;
using System.Globalization;
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
        IsDate1904(wb).Should().BeFalse();
    }

    [Fact]
    public void Excel1900_Option_Leaves_Workbook_In_1900()
    {
        using var wb = Workbook.Create(new WorkbookOptions { DateSystem = DateSystem.Excel1900 });
        IsDate1904(wb).Should().BeFalse();
    }

    [Fact]
    public void Excel1904_Option_Sets_Workbook_Flag()
    {
        using var wb = Workbook.Create(new WorkbookOptions { DateSystem = DateSystem.Excel1904 });
        IsDate1904(wb).Should().BeTrue();
    }

    [Fact]
    public void Excel1904_Shifts_The_Serial_By_The_Epoch_Offset()
    {
        double serial1900, serial1904;

        using (var wb = Workbook.Create(new WorkbookOptions { DateSystem = DateSystem.Excel1900 }))
        {
            wb.AddSheet("S")[1, 1].SetDate(SampleDate);
            serial1900 = SerialOf(wb, "A1");
        }
        using (var wb = Workbook.Create(new WorkbookOptions { DateSystem = DateSystem.Excel1904 }))
        {
            wb.AddSheet("S")[1, 1].SetDate(SampleDate);
            serial1904 = SerialOf(wb, "A1");
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
            FileIsDate1904(path).Should().BeTrue();
            // Reopen with default (1900) options: the file's own epoch must win.
            using (var wb = Workbook.Open(path))
            {
                IsDate1904(wb).Should().BeTrue();
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
                IsDate1904(wb).Should().BeTrue();
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
                wb.Save(path);
            }
            // Streaming Save is single-shot, so the epoch flag is asserted
            // on the saved file rather than mid-write.
            FileIsDate1904(path).Should().BeTrue();
            using (var wb = Workbook.Open(path))
            {
                IsDate1904(wb).Should().BeTrue();
                wb["S"][1, 1].GetDate().Should().Be(SampleDate);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- helpers ------------------------------------------------------

    private static bool IsDate1904(IWorkbook wb)
        => SavedOoxml.BoolAttr(
            SavedOoxml.WorkbookXml(wb).Root!.Element(SavedOoxml.Main + "workbookPr"),
            "date1904");

    private static bool FileIsDate1904(string path)
        => SavedOoxml.BoolAttr(
            SavedOoxml.PartFromFile(path, "xl/workbook.xml").Root!
                .Element(SavedOoxml.Main + "workbookPr"),
            "date1904");

    /// <summary>The raw persisted serial (cell &lt;v&gt;) of a date cell.</summary>
    private static double SerialOf(IWorkbook wb, string a1Address)
        => double.Parse(
            SavedOoxml.Cell(SavedOoxml.SheetXml(wb), a1Address)!
                .Element(SavedOoxml.Main + "v")!.Value,
            CultureInfo.InvariantCulture);
}
