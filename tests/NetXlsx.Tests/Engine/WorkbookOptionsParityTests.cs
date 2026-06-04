// I-82 engine swap — WorkbookOptions parity (pre-cutover parity slice).
//
// The O-15 cutover-readiness scout found the SDK engine ignoring the
// WorkbookOptions read limits (ReadMaxSheets / ReadMaxUncompressedBytes) and
// write caps (MaxRowsPerSheet / MaxColsPerSheet) that the NPOI engine enforces
// (v1.0-A/B contracts). These tests pin the SDK engine to the same contracts
// the NPOI-engine WorkbookOptionsReadPathTests / WorkbookOptionsWritePathTests
// assert, so the cutover inherits them green.

using System;
using System.IO;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class WorkbookOptionsParityTests
{
    // ---- ReadMaxSheets ----------------------------------------------

    [Fact]
    public void ReadMaxSheets_Allows_Open_When_File_Sheet_Count_Is_Within_Limit()
    {
        var path = WriteTempWorkbookWithSheets(5);
        try
        {
            using var wb = Workbook.Open(path, new WorkbookOptions { ReadMaxSheets = 10 });
            wb.SheetCount.Should().Be(5);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void ReadMaxSheets_Throws_ResourceLimitExceededException_When_Exceeded()
    {
        var path = WriteTempWorkbookWithSheets(5);
        try
        {
            Action open = () => Workbook.Open(path, new WorkbookOptions { ReadMaxSheets = 3 });
            open.Should().Throw<ResourceLimitExceededException>()
                .Where(ex => ex.LimitName == "sheet count"
                          && ex.Limit == 3
                          && ex.Actual == 5);
        }
        finally { File.Delete(path); }
    }

    // ---- ReadMaxUncompressedBytes ----------------------------------

    [Fact]
    public void ReadMaxUncompressedBytes_Allows_Open_When_Total_Is_Within_Limit()
    {
        var path = WriteTempWorkbookWithSheets(1);
        try
        {
            using var wb = Workbook.Open(path, new WorkbookOptions
            {
                ReadMaxUncompressedBytes = 10L * 1024 * 1024,
            });
            wb.SheetCount.Should().Be(1);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void ReadMaxUncompressedBytes_Throws_When_Package_Parts_Exceed_Limit()
    {
        // Even a nearly-empty workbook carries several KB of OPC parts
        // (workbook.xml, styles.xml, …), so a 1 KiB limit reliably trips
        // the check without authoring a giant file.
        var path = WriteTempWorkbookWithSheets(1);
        try
        {
            Action open = () => Workbook.Open(path, new WorkbookOptions
            {
                ReadMaxUncompressedBytes = 1024,
            });
            open.Should().Throw<ResourceLimitExceededException>()
                .Where(ex => ex.LimitName == "uncompressed package size in bytes"
                          && ex.Limit == 1024);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Open_That_Fails_The_Read_Limit_Does_Not_Leak_The_Source_Stream()
    {
        // The limit check fires after the package is buffered; the failed open
        // must dispose its own resources and leave the caller's stream usable.
        var path = WriteTempWorkbookWithSheets(3);
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            Action open = () => Workbook.Open(fs, leaveOpen: true,
                new WorkbookOptions { ReadMaxSheets = 1 });
            open.Should().Throw<ResourceLimitExceededException>();
            fs.CanRead.Should().BeTrue("leaveOpen: true must survive a failed open");
            fs.Position = 0;
            using var wb = Workbook.Open(fs, leaveOpen: true);
            wb.SheetCount.Should().Be(3);
        }
        finally { File.Delete(path); }
    }

    // ---- MaxRowsPerSheet --------------------------------------------

    [Fact]
    public void Default_MaxRowsPerSheet_Allows_Excel_Hard_Cap_And_Rejects_Past_It()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => { _ = sheet[CellAddress.MaxRow, 1].Address; }))
            .Should().NotThrow();
        ((Action)(() => { _ = sheet[CellAddress.MaxRow + 1, 1]; }))
            .Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Configured_MaxRowsPerSheet_Caps_AppendRow_And_Indexer()
    {
        using var wb = Workbook.Create(new WorkbookOptions { MaxRowsPerSheet = 5 });
        var sheet = wb.AddSheet("S");
        for (int i = 0; i < 5; i++) sheet.AppendRow();
        ((Action)(() => sheet.AppendRow())).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*exceed the configured row limit of 5*");

        ((Action)(() => { _ = sheet[6, 1]; })).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*row must be in [1, 5]*");

        ((Action)(() => sheet.Row(6))).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*row index must be in [1, 5]*");
    }

    // ---- MaxColsPerSheet --------------------------------------------

    [Fact]
    public void Configured_MaxColsPerSheet_Caps_Indexer_And_Row_Cell_And_Column_Factory()
    {
        using var wb = Workbook.Create(new WorkbookOptions { MaxColsPerSheet = 3 });
        var sheet = wb.AddSheet("S");
        sheet[1, 3].SetString("ok-at-cap");

        ((Action)(() => { _ = sheet[1, 4]; })).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*column must be in [1, 3]*");

        var row = sheet.AppendRow();
        ((Action)(() => row.Cell(4))).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*column must be in [1, 3]*");

        ((Action)(() => sheet.Column(4))).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*column index must be in [1, 3]*");
    }

    [Fact]
    public void Configured_Caps_Apply_To_Range_Coordinates()
    {
        // The NPOI engine validates Range() corners through the same
        // effective-cap gate as the indexer; the SDK engine must agree.
        using var wb = Workbook.Create(new WorkbookOptions { MaxRowsPerSheet = 5, MaxColsPerSheet = 3 });
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet.Range(1, 1, 5, 3))).Should().NotThrow();
        ((Action)(() => sheet.Range(1, 1, 6, 3))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => sheet.Range(1, 1, 5, 4))).Should().Throw<ArgumentOutOfRangeException>();
    }

    // ---- helpers ----------------------------------------------------

    private static string WriteTempWorkbookWithSheets(int count)
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-opts-{Guid.NewGuid():N}.xlsx");
        using var wb = Workbook.Create();
        for (int i = 0; i < count; i++)
        {
            var sheet = wb.AddSheet($"S{i}");
            sheet[1, 1].SetString($"row1-of-sheet{i}");
        }
        wb.Save(path);
        return path;
    }
}
