// I-82 engine swap — streaming write (slice 9) conformance.
//
// Mirrors the NPOI-engine StreamingWorkbookTests behavioral contract on the
// Open XML SDK engine (Workbook.CreateStreamingOoxml): the IStreaming* type
// split (decision #7 / I-13), the append-only row contract, scalar writes,
// formulas, NumberFormat/Style, Save/SaveAsync round-trips, and disposal.
//
// Adds the two slice-9 oracles (advisor, 2026-06-03):
//   1. Cross-engine streaming differential — the SAME dataset streamed through
//      BOTH streaming engines, both outputs reopened through BOTH random-access
//      engines, all four observations compared (output parity; the APIs
//      themselves cannot be compared — NPOI SXSSF is DOM-shaped, the SDK
//      writer is forward-only).
//   2. Forward-only / bounded-memory guard — rows evicted from the row-access
//      window reject writes. If buffering ever silently re-grows to fake
//      random access, the eviction throw disappears and this test fails.
//
// SDK-specific contract (documented divergences, design.md I-82 streaming
// slice): Save is single-shot and fails loud on reuse (NPOI leaks
// ObjectDisposedException from writer internals); writes after Save or past
// the window fail loud (NPOI silently discards them); the NPOI escape hatches
// throw NotSupportedException.

using System;
using System.IO;
using System.Threading.Tasks;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class StreamingEngineTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-stream-{Guid.NewGuid():N}.xlsx");

    // ---- Entry point + lifecycle ------------------------------------

    [Fact]
    public void CreateStreamingOoxml_Returns_Distinct_Type_From_CreateOoxml()
    {
        using var sw = Workbook.CreateStreamingOoxml();
        sw.Should().BeAssignableTo<IStreamingWorkbook>();
        sw.Should().NotBeAssignableTo<IWorkbook>(
            "streaming and random-access are deliberately separate types (decision #7)");
    }

    [Fact]
    public void Double_Dispose_Is_Safe()
    {
        var sw = Workbook.CreateStreamingOoxml();
        sw.AddSheet("S").AppendRow().Set(1, "unsaved");
        sw.Dispose();
        Action again = () => sw.Dispose();
        again.Should().NotThrow();
    }

    // (IStreamingWorkbook/IStreamingSheet.Underlying were REMOVED at v2.0.0 —
    // the streaming engine has no live document to expose; see I-82.)

    // ---- AddSheet ----------------------------------------------------

    [Fact]
    public void AddSheet_Returns_Sheet_With_Same_Name()
    {
        using var sw = Workbook.CreateStreamingOoxml();
        var sheet = sw.AddSheet("Data");
        sheet.Name.Should().Be("Data");
        sheet.Workbook.Should().BeSameAs(sw);
    }

    [Fact]
    public void AddSheet_Rejects_Duplicate_Names_Case_Insensitive()
    {
        using var sw = Workbook.CreateStreamingOoxml();
        sw.AddSheet("Data");
        ((Action)(() => sw.AddSheet("DATA"))).Should().Throw<SheetNameException>();
    }

    [Fact]
    public void AddSheet_Null_Name_Throws_ArgumentNullException()
    {
        using var sw = Workbook.CreateStreamingOoxml();
        ((Action)(() => sw.AddSheet(null!))).Should().Throw<ArgumentNullException>();
    }

    // ---- Append-only contract ---------------------------------------

    [Fact]
    public void AppendRow_Starts_At_Row_1_On_Empty_Sheet()
    {
        using var sw = Workbook.CreateStreamingOoxml();
        sw.AddSheet("S").AppendRow().Index.Should().Be(1);
    }

    [Fact]
    public void AppendRow_Increments_Index_Each_Call()
    {
        using var sw = Workbook.CreateStreamingOoxml();
        var sheet = sw.AddSheet("S");
        sheet.AppendRow().Index.Should().Be(1);
        sheet.AppendRow().Index.Should().Be(2);
        sheet.AppendRow().Index.Should().Be(3);
    }

    [Fact]
    public void AppendRow_Explicit_Index_Skips_Forward()
    {
        using var sw = Workbook.CreateStreamingOoxml();
        var sheet = sw.AddSheet("S");
        sheet.AppendRow();             // row 1
        var row5 = sheet.AppendRow(5); // skip to row 5
        row5.Index.Should().Be(5);
        sheet.AppendRow().Index.Should().Be(6);
    }

    [Fact]
    public void AppendRow_Explicit_Index_Cannot_Revisit_Earlier_Row()
    {
        using var sw = Workbook.CreateStreamingOoxml();
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
        using var sw = Workbook.CreateStreamingOoxml();
        var sheet = sw.AddSheet("S");
        ((Action)(() => sheet.AppendRow(0))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => sheet.AppendRow(CellAddress.MaxRow + 1)))
            .Should().Throw<ArgumentOutOfRangeException>();
    }

    // ---- Cell-level write -------------------------------------------

    [Fact]
    public void Row_Set_Overloads_Write_Each_Scalar_Type()
    {
        var path = TempXlsxPath();
        try
        {
            using (var sw = Workbook.CreateStreamingOoxml())
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
            using (var wb = Workbook.OpenOoxml(path))
            {
                var sheet = wb["S"];
                sheet["A1"].GetString().Should().Be("hello");
                sheet["B1"].GetNumber().Should().Be(3.14);
                sheet["C1"].GetNumber().Should().Be(99.0);
                sheet["D1"].GetNumber().Should().Be(42.0);
                sheet["E1"].GetNumber().Should().Be(100.0);
                sheet["F1"].GetBool().Should().Be(true);
                sheet["G1"].Kind.Should().Be(CellKind.Date);
                sheet["G1"].GetDate()!.Value.Date.Should().Be(new DateTime(2026, 5, 16).Date);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Row_Letter_Indexer_Resolves_Column()
    {
        using var sw = Workbook.CreateStreamingOoxml();
        var row = sw.AddSheet("S").AppendRow();
        row["B"].SetString("at-b");
        row[3].SetString("at-c");
        row.Cell(1).Address.Should().Be("A1");
        row["B"].Address.Should().Be("B1");
        row[3].Address.Should().Be("C1");
    }

    [Fact]
    public void Cell_Bounds_Are_Validated()
    {
        using var sw = Workbook.CreateStreamingOoxml();
        var row = sw.AddSheet("S").AppendRow();
        ((Action)(() => row.Cell(0))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => row.Cell(CellAddress.MaxColumn + 1))).Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Cell_Kind_Tracks_Buffered_Writes()
    {
        using var sw = Workbook.CreateStreamingOoxml();
        var row = sw.AddSheet("S").AppendRow();
        row.Cell(1).Kind.Should().Be(CellKind.Empty);
        row.Cell(1).SetString("x");
        row.Cell(1).Kind.Should().Be(CellKind.String);
        row.Cell(2).SetNumber(1.5);
        row.Cell(2).Kind.Should().Be(CellKind.Number);
        row.Cell(3).SetBool(true);
        row.Cell(3).Kind.Should().Be(CellKind.Bool);
        row.Cell(4).SetDate(new DateTime(2026, 1, 1));
        row.Cell(4).Kind.Should().Be(CellKind.Date);
        row.Cell(5).SetFormula("=A1");
        row.Cell(5).Kind.Should().Be(CellKind.Formula);
    }

    [Fact]
    public void SetString_Over_MaxCellTextLength_Throws()
    {
        using var sw = Workbook.CreateStreamingOoxml(new StreamingOptions { MaxCellTextLength = 8 });
        var row = sw.AddSheet("S").AppendRow();
        ((Action)(() => row.Cell(1).SetString("123456789")))
            .Should().Throw<ResourceLimitExceededException>();
    }

    // ---- Save + round-trip ------------------------------------------

    [Fact]
    public async Task SaveAsync_Produces_Readable_Xlsx()
    {
        var path = TempXlsxPath();
        try
        {
            await using (var sw = Workbook.CreateStreamingOoxml())
            {
                var sheet = sw.AddSheet("S");
                for (int i = 1; i <= 50; i++)
                    sheet.AppendRow().Set(1, $"row {i}").Set(2, i);
                await sw.SaveAsync(path);
            }

            File.Exists(path).Should().BeTrue();
            new FileInfo(path).Length.Should().BeGreaterThan(1000);

            using var wb = Workbook.OpenOoxml(path);
            wb["S"]["A1"].GetString().Should().Be("row 1");
            wb["S"]["B50"].GetNumber().Should().Be(50.0);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Save_To_Stream_Honors_LeaveOpen()
    {
        using var kept = new MemoryStream();
        using (var sw = Workbook.CreateStreamingOoxml())
        {
            sw.AddSheet("S").AppendRow().Set(1, "via-stream");
            sw.Save(kept, leaveOpen: true);
        }
        kept.CanWrite.Should().BeTrue("leaveOpen: true must not dispose the caller's stream");
        kept.Position = 0;
        using var wb = Workbook.OpenOoxml(kept);
        wb["S"]["A1"].GetString().Should().Be("via-stream");
    }

    [Fact]
    public void Save_To_Stream_Disposes_When_LeaveOpen_False()
    {
        var stream = new MemoryStream();
        using (var sw = Workbook.CreateStreamingOoxml())
        {
            sw.AddSheet("S").AppendRow().Set(1, "x");
            sw.Save(stream, leaveOpen: false);
        }
        stream.CanWrite.Should().BeFalse("leaveOpen: false disposes the caller's stream");
    }

    [Fact]
    public void CompressTempFiles_Round_Trips_Identically()
    {
        var path = TempXlsxPath();
        try
        {
            using (var sw = Workbook.CreateStreamingOoxml(
                new StreamingOptions { CompressTempFiles = true, RowAccessWindowSize = 4 }))
            {
                var sheet = sw.AddSheet("S");
                for (int i = 1; i <= 100; i++)
                    sheet.AppendRow().Set(1, $"row {i}").Set(2, i);
                sw.Save(path);
            }
            using var wb = Workbook.OpenOoxml(path);
            wb["S"]["A1"].GetString().Should().Be("row 1");
            wb["S"]["B100"].GetNumber().Should().Be(100.0);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Excel1904_Option_Round_Trips_The_Date()
    {
        var path = TempXlsxPath();
        try
        {
            var date = new DateTime(2026, 5, 16);
            using (var sw = Workbook.CreateStreamingOoxml(
                new StreamingOptions { DateSystem = DateSystem.Excel1904 }))
            {
                sw.AddSheet("S").AppendRow().Set(1, date);
                sw.Save(path);
            }
            using var wb = Workbook.OpenOoxml(path);
            wb["S"]["A1"].Kind.Should().Be(CellKind.Date);
            wb["S"]["A1"].GetDate()!.Value.Should().Be(date,
                "the 1904 serial must be interpreted against the 1904 epoch the file declares");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Single-shot Save (fail-loud divergences) --------------------

    [Fact]
    public void Second_Save_Throws_InvalidOperation()
    {
        var p1 = TempXlsxPath();
        var p2 = TempXlsxPath();
        try
        {
            using var sw = Workbook.CreateStreamingOoxml();
            sw.AddSheet("S").AppendRow().Set(1, "one");
            sw.Save(p1);
            ((Action)(() => sw.Save(p2))).Should().Throw<InvalidOperationException>()
                .WithMessage("*single-shot*",
                    "NPOI leaks ObjectDisposedException from writer internals here; " +
                    "the SDK engine fails loud with an honest message");
        }
        finally
        {
            if (File.Exists(p1)) File.Delete(p1);
            if (File.Exists(p2)) File.Delete(p2);
        }
    }

    [Fact]
    public void Writes_After_Save_Fail_Loud()
    {
        var path = TempXlsxPath();
        try
        {
            using var sw = Workbook.CreateStreamingOoxml();
            var sheet = sw.AddSheet("S");
            var row = sheet.AppendRow();
            row.Set(1, "one");
            sw.Save(path);
            // NPOI silently accepts all of these and loses the data.
            ((Action)(() => sheet.AppendRow())).Should().Throw<InvalidOperationException>();
            ((Action)(() => row.Set(1, "lost"))).Should().Throw<InvalidOperationException>();
            ((Action)(() => sw.AddSheet("T"))).Should().Throw<InvalidOperationException>();
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Failed_Save_To_Bad_Path_Does_Not_Burn_The_Single_Shot()
    {
        var path = TempXlsxPath();
        try
        {
            using var sw = Workbook.CreateStreamingOoxml();
            sw.AddSheet("S").AppendRow().Set(1, "kept");
            var bad = Path.Combine(Path.GetTempPath(), $"netxlsx-no-such-dir-{Guid.NewGuid():N}", "out.xlsx");
            ((Action)(() => sw.Save(bad))).Should().Throw<DirectoryNotFoundException>();
            sw.Save(path); // the failed attempt must not have finalized the sheets
            using var wb = Workbook.OpenOoxml(path);
            wb["S"]["A1"].GetString().Should().Be("kept");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Forward-only / bounded-memory guard (slice-9 oracle 2) ------

    [Fact]
    public void Row_Evicted_From_The_Window_Rejects_Writes()
    {
        using var sw = Workbook.CreateStreamingOoxml(new StreamingOptions { RowAccessWindowSize = 2 });
        var sheet = sw.AddSheet("S");
        var row1 = sheet.AppendRow();
        row1.Set(1, "one");
        sheet.AppendRow().Set(1, "two");
        sheet.AppendRow().Set(1, "three"); // evicts row 1 past the 2-row window

        // The forward-only property: row 1 is on disk, not in memory. If
        // internal buffering ever silently re-grows to fake random access,
        // this write would succeed and the guard fails.
        ((Action)(() => row1.Set(1, "rewrite"))).Should().Throw<InvalidOperationException>()
            .WithMessage("*flushed*");
        ((Action)(() => _ = row1.Cell(1).Kind)).Should().Throw<InvalidOperationException>();
    }

    [Fact]
    public void Explicit_Flush_Releases_The_Window_And_Appending_Continues()
    {
        var path = TempXlsxPath();
        try
        {
            using (var sw = Workbook.CreateStreamingOoxml())
            {
                var sheet = sw.AddSheet("S");
                var row = sheet.AppendRow();
                row.Set(1, "x");
                row.Flush(); // SXSSFSheet.FlushRows() semantics: the whole window flushes
                ((Action)(() => row.Set(2, "late"))).Should().Throw<InvalidOperationException>()
                    .WithMessage("*flushed*");
                sheet.AppendRow().Set(1, "after-flush"); // appending forward stays open
                sw.Save(path);
            }
            using var wb = Workbook.OpenOoxml(path);
            wb["S"]["A1"].GetString().Should().Be("x");
            wb["S"]["A2"].GetString().Should().Be("after-flush");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Interleaved_Writes_Across_Sheets_Persist_Independently()
    {
        // Each sheet owns its own forward-only writer + window, so callers can
        // alternate between sheets mid-stream (the SXSSF usage pattern).
        var path = TempXlsxPath();
        try
        {
            using (var sw = Workbook.CreateStreamingOoxml(new StreamingOptions { RowAccessWindowSize = 2 }))
            {
                var a = sw.AddSheet("A");
                var b = sw.AddSheet("B");
                for (int i = 1; i <= 10; i++)
                {
                    a.AppendRow().Set(1, $"a{i}");
                    b.AppendRow().Set(1, $"b{i}");
                }
                sw.Save(path);
            }
            using var wb = Workbook.OpenOoxml(path);
            wb["A"]["A1"].GetString().Should().Be("a1");
            wb["A"]["A10"].GetString().Should().Be("a10");
            wb["B"]["A1"].GetString().Should().Be("b1");
            wb["B"]["A10"].GetString().Should().Be("b10");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Windowed_Write_Of_Many_Rows_Persists_All_Of_Them()
    {
        var path = TempXlsxPath();
        try
        {
            using (var sw = Workbook.CreateStreamingOoxml(new StreamingOptions { RowAccessWindowSize = 16 }))
            {
                var sheet = sw.AddSheet("S");
                for (int i = 1; i <= 10_000; i++)
                    sheet.AppendRow().Set(1, i).Set(2, $"row {i}");
                sw.Save(path);
            }
            using var wb = Workbook.OpenOoxml(path);
            wb["S"]["A1"].GetNumber().Should().Be(1.0);
            wb["S"]["B1"].GetString().Should().Be("row 1");
            wb["S"]["A10000"].GetNumber().Should().Be(10_000.0);
            wb["S"]["B10000"].GetString().Should().Be("row 10000");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Disposed-after-Dispose throws on every member ----------------

    [Fact]
    public void Members_Throw_ObjectDisposed_After_Dispose()
    {
        var sw = Workbook.CreateStreamingOoxml();
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
        var path = TempXlsxPath();
        try
        {
            using (var sw = Workbook.CreateStreamingOoxml())
            {
                var row = sw.AddSheet("S").AppendRow();
                row.Set(1, 10).Set(2, 20);
                row.Cell(3).SetFormula("=A1+B1");
                sw.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                wb["S"]["C1"].Kind.Should().Be(CellKind.Formula);
                wb["S"]["C1"].GetFormula().Should().Be("=A1+B1");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Cell_SetFormula_Empty_Or_Broken_Throws()
    {
        using var sw = Workbook.CreateStreamingOoxml();
        var row = sw.AddSheet("S").AppendRow();
        ((Action)(() => row.Cell(1).SetFormula(""))).Should().Throw<FormulaException>();
        ((Action)(() => row.Cell(1).SetFormula(null!))).Should().Throw<ArgumentNullException>();
        // The cross-engine-pinned structural rejection (same as the DOM engines).
        ((Action)(() => row.Cell(1).SetFormula("=SUM("))).Should().Throw<FormulaException>();
    }

    // ---- NumberFormat / Style ---------------------------------------

    [Fact]
    public void NumberFormat_Survives_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            using (var sw = Workbook.CreateStreamingOoxml())
            {
                var row = sw.AddSheet("S").AppendRow();
                row.Cell(1).SetNumber(1234.56);
                row.Cell(1).NumberFormat(NumberFormats.Currency);
                sw.Save(path);
            }
            using var wb = Workbook.OpenOoxml(path);
            wb["S"]["A1"].GetStyle().NumberFormat.Should().Be(NumberFormats.Currency);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Style_Is_Pool_Deduped_And_Survives_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            var style = new CellStyle { Bold = true, Background = Color.FromRgb(0xFF, 0xEE, 0x00) };
            using (var sw = Workbook.CreateStreamingOoxml())
            {
                var sheet = sw.AddSheet("S");
                for (int i = 1; i <= 20; i++)
                {
                    var row = sheet.AppendRow();
                    row.Cell(1).SetString($"r{i}");
                    row.Cell(1).Style(style);
                }
                sw.Save(path);
            }
            using var wb = Workbook.OpenOoxml(path);
            wb["S"]["A1"].GetStyle().Bold.Should().BeTrue();
            wb["S"]["A20"].GetStyle().Bold.Should().BeTrue();
            wb["S"]["A1"].GetStyle().Background.Should().Be(Color.FromRgb(0xFF, 0xEE, 0x00));
            // One shared style object across 20 cells -> one cellXfs entry
            // (pool dedup, decision #4 / #29). Index 0 is the default xf.
            wb.Underlying.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
                .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.CellFormats>()!
                .Count!.Value.Should().Be(2u);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Style_Merge_Preserves_NumberFormat_Like_The_Npoi_Engine()
    {
        var path = TempXlsxPath();
        try
        {
            using (var sw = Workbook.CreateStreamingOoxml())
            {
                var row = sw.AddSheet("S").AppendRow();
                row.Cell(1).SetNumber(1234.56);
                row.Cell(1).NumberFormat("#,##0.00");
                // A later Style() without NumberFormat must keep the format —
                // the one axis the streaming merge reads back (SxssfCell parity).
                row.Cell(1).Style(new CellStyle { Bold = true });
                sw.Save(path);
            }
            using var wb = Workbook.OpenOoxml(path);
            wb["S"]["A1"].GetStyle().NumberFormat.Should().Be("#,##0.00");
            wb["S"]["A1"].GetStyle().Bold.Should().BeTrue();
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Cross-engine streaming differential (slice-9 oracle 1) -------

    private sealed record StreamProjection(
        string A1, double? B1, double? C1, bool? D1, DateTime? E1, CellKind E1Kind,
        string? F1Formula, CellKind F1Kind, string? Fmt, bool? Bold, string A5, string SecondA1);

    private static void BuildDataset(IStreamingWorkbook sw)
    {
        var s1 = sw.AddSheet("Data");
        var r1 = s1.AppendRow();
        r1.Set(1, "hello").Set(2, 3.14).Set(3, 42).Set(4, true).Set(5, new DateTime(2026, 5, 16));
        r1.Cell(6).SetFormula("=A1&B1");
        var r2 = s1.AppendRow();
        r2.Cell(1).SetNumber(1234.56);
        r2.Cell(1).NumberFormat("#,##0.00");
        r2.Cell(2).SetString("styled");
        r2.Cell(2).Style(new CellStyle { Bold = true });
        s1.AppendRow(5).Set(1, "row5"); // skipped rows stay sparse
        sw.AddSheet("Second").AppendRow().Set(1, "s2");
    }

    private static StreamProjection ReadDataset(IWorkbook wb)
    {
        var s = wb["Data"];
        return new StreamProjection(
            s["A1"].GetString(), s["B1"].GetNumber(), s["C1"].GetNumber(), s["D1"].GetBool(),
            s["E1"].GetDate(), s["E1"].Kind,
            s["F1"].GetFormula(), s["F1"].Kind,
            s["A2"].GetStyle().NumberFormat, s["B2"].GetStyle().Bold,
            s["A5"].GetString(), wb["Second"]["A1"].GetString());
    }

    [Fact]
    public void Streamed_Output_Round_Trips_The_Dataset_Exactly()
    {
        // Was the slice-9 cross-engine 4-way differential. The cross-engine
        // de-risk mission completed at the v2.0.0 cutover (the re-scout is the
        // evidence; A1 disposition (b)) — the observation is now pinned as
        // LITERALS, so the round trip asserts against expected values rather
        // than a second engine.
        var path = TempXlsxPath();
        try
        {
            using (var sw = Workbook.CreateStreaming()) { BuildDataset(sw); sw.Save(path); }

            using var wb = Workbook.Open(path);
            ReadDataset(wb).Should().Be(new StreamProjection(
                "hello", 3.14, 42, true, new DateTime(2026, 5, 16), CellKind.Date,
                "=A1&B1", CellKind.Formula, "#,##0.00", true, "row5", "s2"));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ---- Schema gate ---------------------------------------------------

    [Fact]
    public void Streamed_Output_Is_Schema_Valid()
    {
        var path = TempXlsxPath();
        try
        {
            using (var sw = Workbook.CreateStreamingOoxml(
                new StreamingOptions { RowAccessWindowSize = 2, DateSystem = DateSystem.Excel1904 }))
            {
                BuildDataset(sw);
                sw.AddSheet("Empty"); // a sheet with no rows must also validate
                sw.Save(path);
            }
            using var wb = Workbook.OpenOoxml(path);
            OpenXmlValidationGate.AssertValid(wb);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
