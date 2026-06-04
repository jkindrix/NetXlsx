// Coverage for the v0.9 streaming-write API (IStreamingWorkbook).

using System;
using System.IO;
using System.Threading.Tasks;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class StreamingWorkbookTests
{
    // ---- Entry point + lifecycle ------------------------------------

    [Fact]
    public void CreateStreaming_Returns_Distinct_Type_From_Create()
    {
        using var sw = Workbook.CreateStreaming();
        sw.Should().BeAssignableTo<IStreamingWorkbook>();
        sw.Should().NotBeAssignableTo<IWorkbook>(
            "streaming and random-access are deliberately separate types (decision #7)");
    }

    // The window-size plumbing is asserted BEHAVIORALLY since v2.0.0 (the
    // SDK engine's window is internal; the old test reached through the
    // removed NPOI hatch): a row evicted from the access window rejects
    // writes, so where the eviction boundary falls proves which window
    // size the engine is running.

    [Fact]
    public void CreateStreaming_Honors_Explicit_Window_Size()
    {
        using var sw = Workbook.CreateStreaming(new StreamingOptions { RowAccessWindowSize = 2 });
        var sheet = sw.AddSheet("S");
        var r1 = sheet.AppendRow();
        r1.Set(1, "one");
        sheet.AppendRow().Set(1, "two");
        sheet.AppendRow().Set(1, "three"); // evicts row 1 (window = 2)

        ((Action)(() => r1.Cell(2).SetString("late")))
            .Should().Throw<InvalidOperationException>(
                "row 1 left the 2-row access window, so the configured size is plumbed through");
    }

    [Fact]
    public void CreateStreaming_Default_Window_Keeps_Early_Rows_Writable()
    {
        // Default RowAccessWindowSize is 100 (design §6.1): after 50 appends
        // row 1 is still inside the window and accepts writes.
        using var sw = Workbook.CreateStreaming();
        var sheet = sw.AddSheet("S");
        var r1 = sheet.AppendRow();
        r1.Set(1, "one");
        for (int i = 2; i <= 50; i++) sheet.AppendRow().Set(1, $"row {i}");

        ((Action)(() => r1.Cell(2).SetString("still writable"))).Should().NotThrow();
    }

    [Fact]
    public void Double_Dispose_Is_Safe()
    {
        var sw = Workbook.CreateStreaming();
        sw.Dispose();
        Action again = () => sw.Dispose();
        again.Should().NotThrow();
    }

    // ---- AddSheet ----------------------------------------------------

    [Fact]
    public void AddSheet_Returns_Sheet_With_Same_Name()
    {
        using var sw = Workbook.CreateStreaming();
        var sheet = sw.AddSheet("Data");
        sheet.Name.Should().Be("Data");
        sheet.Workbook.Should().BeSameAs(sw);
    }

    [Fact]
    public void AddSheet_Rejects_Duplicate_Names_Case_Insensitive()
    {
        using var sw = Workbook.CreateStreaming();
        sw.AddSheet("Data");
        ((Action)(() => sw.AddSheet("DATA"))).Should().Throw<SheetNameException>();
    }

    [Fact]
    public void AddSheet_Null_Name_Throws_ArgumentNullException()
    {
        using var sw = Workbook.CreateStreaming();
        ((Action)(() => sw.AddSheet(null!))).Should().Throw<ArgumentNullException>();
    }

    // ---- Append-only contract ---------------------------------------

    [Fact]
    public void AppendRow_Starts_At_Row_1_On_Empty_Sheet()
    {
        using var sw = Workbook.CreateStreaming();
        var sheet = sw.AddSheet("S");
        var row = sheet.AppendRow();
        row.Index.Should().Be(1);
    }

    [Fact]
    public void AppendRow_Increments_Index_Each_Call()
    {
        using var sw = Workbook.CreateStreaming();
        var sheet = sw.AddSheet("S");
        sheet.AppendRow().Index.Should().Be(1);
        sheet.AppendRow().Index.Should().Be(2);
        sheet.AppendRow().Index.Should().Be(3);
    }

    [Fact]
    public void AppendRow_Explicit_Index_Skips_Forward()
    {
        using var sw = Workbook.CreateStreaming();
        var sheet = sw.AddSheet("S");
        sheet.AppendRow();             // row 1
        var row5 = sheet.AppendRow(5); // skip to row 5
        row5.Index.Should().Be(5);
        sheet.AppendRow().Index.Should().Be(6);
    }

    [Fact]
    public void AppendRow_Explicit_Index_Cannot_Revisit_Earlier_Row()
    {
        using var sw = Workbook.CreateStreaming();
        var sheet = sw.AddSheet("S");
        sheet.AppendRow();   // row 1
        sheet.AppendRow();   // row 2
        ((Action)(() => sheet.AppendRow(2))).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*append-only*");
        ((Action)(() => sheet.AppendRow(1))).Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void AppendRow_Explicit_Index_Out_Of_Grid_Throws()
    {
        using var sw = Workbook.CreateStreaming();
        var sheet = sw.AddSheet("S");
        ((Action)(() => sheet.AppendRow(0))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => sheet.AppendRow(CellAddress.MaxRow + 1)))
            .Should().Throw<ArgumentOutOfRangeException>();
    }

    // ---- Cell-level write -------------------------------------------

    [Fact]
    public void Row_Set_Overloads_Write_Each_Scalar_Type()
    {
        var path = Path.Combine(Path.GetTempPath(), $"stream-scalars-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var sw = Workbook.CreateStreaming())
            {
                var sheet = sw.AddSheet("S");
                sheet.AppendRow()
                    .Set(1, "hello")
                    .Set(2, 3.14)
                    .Set(3, 99m)
                    .Set(4, 42)
                    .Set(5, 100L)
                    .Set(6, true)
                    .Set(7, new DateTime(2026, 5, 16));
                sw.Save(path);
            }
            // Read via the random-access API to verify content.
            using (var wb = Workbook.Open(path))
            {
                var sheet = wb["S"];
                sheet["A1"].GetString().Should().Be("hello");
                sheet["B1"].GetNumber().Should().Be(3.14);
                sheet["C1"].GetNumber().Should().Be(99.0);
                sheet["D1"].GetNumber().Should().Be(42.0);
                sheet["E1"].GetNumber().Should().Be(100.0);
                sheet["F1"].GetBool().Should().Be(true);
                sheet["G1"].GetDate()!.Value.Date.Should().Be(new DateTime(2026, 5, 16).Date);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Row_Letter_Indexer_Resolves_Column()
    {
        using var sw = Workbook.CreateStreaming();
        var sheet = sw.AddSheet("S");
        var row = sheet.AppendRow();
        row["B"].SetString("at-b");
        row[3].SetString("at-c");
        row.Cell(1).Address.Should().Be("A1");
        row["B"].Address.Should().Be("B1");
        row[3].Address.Should().Be("C1");
    }

    [Fact]
    public void Cell_Bounds_Are_Validated()
    {
        using var sw = Workbook.CreateStreaming();
        var sheet = sw.AddSheet("S");
        var row = sheet.AppendRow();
        ((Action)(() => row.Cell(0))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => row.Cell(CellAddress.MaxColumn + 1))).Should().Throw<ArgumentOutOfRangeException>();
    }

    // ---- Save + round-trip via Workbook.Open ------------------------

    [Fact]
    public async Task SaveAsync_Produces_Readable_Xlsx()
    {
        var path = Path.Combine(Path.GetTempPath(), $"stream-save-{Guid.NewGuid():N}.xlsx");
        try
        {
            await using (var sw = Workbook.CreateStreaming())
            {
                var sheet = sw.AddSheet("S");
                for (int i = 1; i <= 50; i++)
                    sheet.AppendRow().Set(1, $"row {i}").Set(2, i);
                await sw.SaveAsync(path);
            }

            File.Exists(path).Should().BeTrue();
            new FileInfo(path).Length.Should().BeGreaterThan(1000);

            using var wb = Workbook.Open(path);
            wb["S"]["A1"].GetString().Should().Be("row 1");
            wb["S"]["B50"].GetNumber().Should().Be(50.0);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Disposed-after Dispose throws on every member --------------

    [Fact]
    public void Members_Throw_ObjectDisposed_After_Dispose()
    {
        var sw = Workbook.CreateStreaming();
        var sheet = sw.AddSheet("S");
        var row = sheet.AppendRow();
        var cell = row.Cell(1);
        sw.Dispose();

        ((Action)(() => sw.AddSheet("T"))).Should().Throw<ObjectDisposedException>();
        ((Action)(() => sheet.AppendRow())).Should().Throw<ObjectDisposedException>();
        ((Action)(() => row.Cell(2))).Should().Throw<ObjectDisposedException>();
        ((Action)(() => row.Flush())).Should().Throw<ObjectDisposedException>();
        ((Action)(() => cell.SetString("x"))).Should().Throw<ObjectDisposedException>();
        ((Action)(() => cell.NumberFormat("0.00"))).Should().Throw<ObjectDisposedException>();
    }

    // ---- Formula on streaming cell ----------------------------------

    [Fact]
    public void Cell_SetFormula_Stores_Body_Without_Pre_Computation()
    {
        var path = Path.Combine(Path.GetTempPath(), $"stream-formula-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var sw = Workbook.CreateStreaming())
            {
                var sheet = sw.AddSheet("S");
                var row = sheet.AppendRow();
                row.Set(1, 10).Set(2, 20);
                row.Cell(3).SetFormula("=A1+B1");
                sw.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["S"]["C1"].Kind.Should().Be(CellKind.Formula);
                wb["S"]["C1"].GetFormula().Should().Be("=A1+B1");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Cell_SetFormula_Empty_Throws()
    {
        using var sw = Workbook.CreateStreaming();
        var row = sw.AddSheet("S").AppendRow();
        ((Action)(() => row.Cell(1).SetFormula(""))).Should().Throw<FormulaException>();
        ((Action)(() => row.Cell(1).SetFormula(null!))).Should().Throw<ArgumentNullException>();
    }

    // ---- NumberFormat / Style ---------------------------------------

    [Fact]
    public void NumberFormat_Survives_Save_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"stream-fmt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var sw = Workbook.CreateStreaming())
            {
                var sheet = sw.AddSheet("S");
                var row = sheet.AppendRow();
                row.Cell(1).SetNumber(1234.56);
                row.Cell(1).NumberFormat(NumberFormats.Currency);
                sw.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["S"]["A1"].GetStyle().NumberFormat.Should().Be(NumberFormats.Currency);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
