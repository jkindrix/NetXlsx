// End-to-end round-trip: Create -> AddSheet -> SetX -> SaveAsync ->
// OpenAsync -> GetX. Exercises the v0.2.0 vertical slice.

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using FluentAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class WorkbookRoundTripTests
{
    [Fact]
    public async Task Round_Trips_String_Number_Bool_Across_Save_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-roundtrip-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("Data");
                sheet["A1"].SetString("Hello");
                sheet["B1"].SetNumber(42.5);
                sheet["C1"].SetNumber(1234.56m);
                sheet["D1"].SetBool(true);
                await wb.SaveAsync(path);
            }

            using (var wb = await Workbook.OpenAsync(path))
            {
                wb.SheetCount.Should().Be(1);
                var sheet = wb["Data"];
                sheet.Name.Should().Be("Data");

                sheet["A1"].Kind.Should().Be(CellKind.String);
                sheet["A1"].GetString().Should().Be("Hello");

                sheet["B1"].Kind.Should().Be(CellKind.Number);
                sheet["B1"].GetNumber().Should().Be(42.5);

                sheet["C1"].Kind.Should().Be(CellKind.Number);
                sheet["C1"].GetNumber().Should().Be(1234.56);

                sheet["D1"].Kind.Should().Be(CellKind.Bool);
                sheet["D1"].GetBool().Should().Be(true);
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Cell_Indexer_Materializes_Empty_Cells_On_Access()
    {
        // Decision #40: accessing an unwritten cell auto-materializes as blank.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var cell = sheet["Z99"];
        cell.Kind.Should().Be(CellKind.Empty);
        cell.RowIndex.Should().Be(99);
        cell.ColumnIndex.Should().Be(26);
        cell.Address.Should().Be("Z99");
        cell.GetString().Should().Be("");
        cell.GetNumber().Should().BeNull();
        cell.GetBool().Should().BeNull();
    }

    [Fact]
    public void Cell_Address_Is_Canonical_On_Roundtrip()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["$a$1"].Address.Should().Be("A1");
        sheet["aa5"].Address.Should().Be("AA5");
    }

    [Fact]
    public void AddSheet_Throws_On_Invalid_Names()
    {
        using var wb = Workbook.Create();

        Action tooLong = () => wb.AddSheet(new string('x', 32));
        tooLong.Should().Throw<SheetNameException>().WithMessage("*exceeds 31 characters*");

        Action invalidChar = () => wb.AddSheet("Bad/Name");
        invalidChar.Should().Throw<SheetNameException>().WithMessage("*invalid character*");

        Action empty = () => wb.AddSheet("");
        empty.Should().Throw<SheetNameException>().WithMessage("*empty*");
    }

    [Fact]
    public void AddSheet_Throws_On_Duplicate_Name_Case_Insensitive()
    {
        // Decision #41: case-insensitive uniqueness.
        using var wb = Workbook.Create();
        wb.AddSheet("Data");

        Action duplicate = () => wb.AddSheet("DATA");
        duplicate.Should().Throw<SheetNameException>().WithMessage("*already exists*");
    }

    [Fact]
    public void Use_After_Dispose_Throws_ObjectDisposedException()
    {
        // Decision #42.
        var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Dispose();

        Action addSheet = () => wb.AddSheet("T");
        addSheet.Should().Throw<ObjectDisposedException>();

        Action indexerSheets = () => { var _ = wb.SheetCount; };
        indexerSheets.Should().Throw<ObjectDisposedException>();
    }

    [Fact]
    public void Double_Dispose_Is_Safe()
    {
        // Decision #42.
        var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Dispose();
        Action again = () => wb.Dispose();
        again.Should().NotThrow();
    }

    [Fact]
    public async Task Concurrent_AddSheet_Throws_InvalidOperationException()
    {
        // Decision #43: workbooks are not thread-safe, but we detect
        // concurrent mutation rather than letting it silently corrupt.
        // Drive two threads racing on AddSheet and assert at least one
        // sees InvalidOperationException.
        //
        // Previously this test flaked occasionally (one thread would
        // finish before the other entered its loop). The Barrier forces
        // both threads to begin their tight loops at the same instant,
        // and the iteration count is bumped 5x so a single scheduler
        // hiccup cannot make the race vanish entirely on a one-shot run.
        using var wb = Workbook.Create();

        const int iterations = 1_000;
        var caughtConcurrent = 0;
        using var barrier = new Barrier(2);

        await Task.WhenAll(
            Task.Run(() =>
            {
                barrier.SignalAndWait();
                for (int i = 0; i < iterations; i++)
                {
                    try { wb.AddSheet($"A_{i}"); }
                    catch (InvalidOperationException) { Interlocked.Increment(ref caughtConcurrent); }
                    catch (SheetNameException) { /* duplicate from racing thread — ignore */ }
                }
            }),
            Task.Run(() =>
            {
                barrier.SignalAndWait();
                for (int i = 0; i < iterations; i++)
                {
                    try { wb.AddSheet($"B_{i}"); }
                    catch (InvalidOperationException) { Interlocked.Increment(ref caughtConcurrent); }
                    catch (SheetNameException) { /* duplicate from racing thread — ignore */ }
                }
            }));

        // Detection is best-effort by design (it's not a lock), but with
        // 2,000 racing mutations starting at exactly the same instant
        // we should observe at least one detected collision.
        caughtConcurrent.Should().BeGreaterThan(0,
            "concurrent AddSheet calls should trigger the reentry-counter detection at least once");
    }

    [Fact]
    public void Open_Rejects_Non_Xlsx_Stream()
    {
        // Decision #51 / I14.
        using var stream = new MemoryStream(new byte[] { 0x00, 0x01, 0x02, 0x03 });
        Action open = () => Workbook.Open(stream);
        open.Should().Throw<MalformedFileException>();
    }

    [Fact]
    public void Open_Rejects_Non_Seekable_Stream()
    {
        // Decision I14: NPOI requires seek.
        using var inner = new MemoryStream();
        using var nonSeekable = new NonSeekableStream(inner);
        Action open = () => Workbook.Open(nonSeekable);
        open.Should().Throw<ArgumentException>().WithMessage("*seekable*");
    }

    [Fact]
    public void Open_Rejects_Stream_Not_At_Position_Zero()
    {
        // Decision #50.
        using var stream = new MemoryStream(new byte[] { 0, 1, 2 });
        stream.Position = 1;
        Action open = () => Workbook.Open(stream);
        open.Should().Throw<ArgumentException>().WithMessage("*positioned at 0*");
    }

    [Fact]
    public void TryGetSheet_Returns_False_For_Missing()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Real");
        wb.TryGetSheet("missing", out _).Should().BeFalse();
        wb.TryGetSheet("Real", out var sheet).Should().BeTrue();
        sheet!.Name.Should().Be("Real");
    }

    [Fact]
    public void IsValidSheetName_And_SanitizeSheetName_Match_Documented_Rules()
    {
        Workbook.IsValidSheetName("OK").Should().BeTrue();
        Workbook.IsValidSheetName("").Should().BeFalse();
        Workbook.IsValidSheetName(new string('x', 32)).Should().BeFalse();
        Workbook.IsValidSheetName("Has/Slash").Should().BeFalse();

        Workbook.SanitizeSheetName("Has/Slash").Should().Be("Has_Slash");
        Workbook.SanitizeSheetName("").Should().Be("Sheet");
        Workbook.SanitizeSheetName(new string('x', 40)).Length.Should().Be(31);
    }

    private sealed class NonSeekableStream : Stream
    {
        private readonly Stream _inner;
        public NonSeekableStream(Stream inner) { _inner = inner; }
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => _inner.CanWrite;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
    }
}
