// Coverage for ISheet.LastRowNumber (decision I-85): the 1-based index of
// the last row containing at least one cell, 0 when the sheet has no cells.
// Added in the pre-cutover parity slice so the source generator's ReadRows
// body can stop reaching through sheet.Underlying.LastRowNum (NPOI-typed,
// 0-based) — see the I-85 row in docs/design.md.

using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class LastRowNumberTests
{
    [Fact]
    public void Empty_Sheet_Reports_Zero()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S").LastRowNumber.Should().Be(0);
    }

    [Fact]
    public void Reports_The_Last_Row_That_Contains_A_Cell()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet[1, 1].SetString("a");
        sheet[5, 3].SetNumber(42);
        sheet.LastRowNumber.Should().Be(5);
    }

    [Fact]
    public void A_Materialized_Row_Without_Cells_Does_Not_Count()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet[2, 1].SetString("data");
        _ = sheet.Row(10);          // materializes an empty row element
        _ = sheet.AppendRow();      // likewise
        sheet.LastRowNumber.Should().Be(2,
            "rows materialized via Row(int)/AppendRow() hold no cells and must not count (I-85)");
    }

    [Fact]
    public void A_Cleared_Cell_Still_Counts()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet[1, 1].SetString("a");
        sheet[7, 1].SetString("b");
        sheet[7, 1].Clear();
        sheet.LastRowNumber.Should().Be(7,
            "Clear() blanks the cell without removing it, so the row still contains a cell (I-85)");
    }

    [Fact]
    public void Survives_A_Save_Open_Round_Trip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"lastrow-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("S");
                sheet[3, 2].SetString("x");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["S"].LastRowNumber.Should().Be(3);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Throws_ObjectDisposedException_After_Dispose()
    {
        var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        wb.Dispose();
        ((Action)(() => { _ = sheet.LastRowNumber; }))
            .Should().Throw<ObjectDisposedException>();
    }
}
