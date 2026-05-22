// Coverage for the v1.1 AutoFilter slice: ISheet.SetAutoFilter,
// ClearAutoFilter, HasAutoFilter, AutoFilterRange.

using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class AutoFilterTests
{
    [Fact]
    public void Sheet_Has_No_AutoFilter_By_Default()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.HasAutoFilter.Should().BeFalse();
        sh.AutoFilterRange.Should().BeNull();
    }

    [Fact]
    public void SetAutoFilter_Records_Range()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:D10");
        sh.HasAutoFilter.Should().BeTrue();
        sh.AutoFilterRange.Should().Be("A1:D10");
    }

    [Fact]
    public void SetAutoFilter_Replaces_Existing()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B5");
        sh.SetAutoFilter("C1:E20");
        sh.AutoFilterRange.Should().Be("C1:E20");
    }

    [Fact]
    public void ClearAutoFilter_Removes_The_Filter()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:D10");
        sh.ClearAutoFilter();
        sh.HasAutoFilter.Should().BeFalse();
        sh.AutoFilterRange.Should().BeNull();
    }

    [Fact]
    public void ClearAutoFilter_On_Empty_Sheet_Is_Safe()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.ClearAutoFilter();
        act.Should().NotThrow();
    }

    [Fact]
    public void SetAutoFilter_Rejects_Null_Range()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.SetAutoFilter(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void SetAutoFilter_Rejects_Invalid_Range()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.SetAutoFilter("notarange");
        act.Should().Throw<InvalidCellAddressException>();
    }

    [Fact]
    public void SetAutoFilter_Accepts_Single_Cell_Range()
    {
        // Single-cell autofilter is permitted (Excel treats it as a
        // 1×1 filterable area). Verify we don't reject it.
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1");
        sh.HasAutoFilter.Should().BeTrue();
        // NPOI's CellRangeAddress.FormatAsString collapses 1×1 to "A1".
        sh.AutoFilterRange.Should().Be("A1");
    }

    [Fact]
    public void AutoFilter_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"af-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh["A1"].SetString("Name");
                sh["B1"].SetString("Score");
                sh["A2"].SetString("Alice"); sh["B2"].SetNumber(90);
                sh["A3"].SetString("Bob");   sh["B3"].SetNumber(85);
                sh.SetAutoFilter("A1:B3");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var sh = wb["S"];
                sh.HasAutoFilter.Should().BeTrue();
                sh.AutoFilterRange.Should().Be("A1:B3");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
