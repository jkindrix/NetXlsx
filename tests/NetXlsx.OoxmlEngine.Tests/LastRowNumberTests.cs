// I-82 engine swap — ISheet.LastRowNumber conformance (decision I-85,
// pre-cutover parity slice).
//
// The SDK-engine half of the I-85 contract (the NPOI-engine half lives in
// tests/NetXlsx.Tests/LastRowNumberTests.cs), plus a cross-engine agreement
// check: the member exists precisely so the source generator's ReadRows body
// can be engine-agnostic, so both engines must report identical values for
// the same content.

using System;
using System.IO;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class LastRowNumberTests
{
    [Fact]
    public void Empty_Sheet_Reports_Zero()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S").LastRowNumber.Should().Be(0);
    }

    [Fact]
    public void Reports_The_Last_Row_That_Contains_A_Cell()
    {
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet[1, 1].SetString("a");
        sheet[5, 3].SetNumber(42);
        sheet.LastRowNumber.Should().Be(5);
    }

    [Fact]
    public void A_Materialized_Row_Without_Cells_Does_Not_Count()
    {
        using var wb = Workbook.CreateOoxml();
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
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet[1, 1].SetString("a");
        sheet[7, 1].SetString("b");
        sheet[7, 1].Clear();
        sheet.LastRowNumber.Should().Be(7,
            "Clear() keeps the <c> node, so the row still contains a cell (I-85)");
    }

    [Fact]
    public void Survives_A_Save_Open_Round_Trip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-lastrow-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var sheet = wb.AddSheet("S");
                sheet[3, 2].SetString("x");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                wb["S"].LastRowNumber.Should().Be(3);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Throws_ObjectDisposedException_After_Dispose()
    {
        var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        wb.Dispose();
        ((Action)(() => { _ = sheet.LastRowNumber; }))
            .Should().Throw<ObjectDisposedException>();
    }

    [Fact]
    public void Engines_Agree_On_Every_Contract_Scenario()
    {
        // Same script through both factories — the values must be identical,
        // because generated ReadRows code runs unchanged on either engine.
        static int[] Script(IWorkbook wb)
        {
            var sheet = wb.AddSheet("S");
            var observations = new int[4];
            observations[0] = sheet.LastRowNumber;            // empty: 0
            sheet[4, 2].SetString("x");
            observations[1] = sheet.LastRowNumber;            // 4
            _ = sheet.Row(9);
            observations[2] = sheet.LastRowNumber;            // still 4
            sheet[6, 1].SetNumber(1);
            sheet[6, 1].Clear();
            observations[3] = sheet.LastRowNumber;            // 6 (cleared cell counts)
            return observations;
        }

        int[] npoi, sdk;
        using (var wb = Workbook.Create()) npoi = Script(wb);
        using (var wb = Workbook.CreateOoxml()) sdk = Script(wb);

        sdk.Should().Equal(npoi, "I-85's contract must be engine-stable by construction");
        sdk.Should().Equal(0, 4, 4, 6);
    }
}
