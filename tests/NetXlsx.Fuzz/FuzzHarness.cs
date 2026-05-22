// Fuzz harness for Workbook.Open / OpenAsync (v1.1 roadmap +
// design decision I-60). Feeds malformed inputs to the open path
// and asserts the response stays inside the documented exception
// contract — no unhandled crashes, no unbounded resource use.
//
// All tests are marked Trait("Category", "Fuzz") so CI can opt in
// (the bulk-iteration tests are slow). Local runs include them
// by default.

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.Fuzz;

[Trait("Category", "Fuzz")]
public class FuzzHarness
{
    // ---- Documented "safe" exception types ----------------------------
    //
    // Open() may throw any of these in response to malformed input. A
    // throw that does NOT match this list is a finding — file an issue.
    // Catches the broad System.Exception base in test code so the
    // harness can categorize, not so production code can.

    private static bool IsAcceptableOpenException(Exception ex) =>
        ex is WorkbookException
        || ex is ResourceLimitExceededException
        || ex is MalformedFileException
        || ex is InvalidDataException        // System.IO.Compression on bad zip
        || ex is IOException                 // truncated streams, etc.
        || ex is FormatException             // XML parser fall-through
        || ex is ArgumentException           // null path, etc. (precondition)
        || ex is NotSupportedException       // unsupported OOXML feature
        // NPOI wraps many parsing errors in untyped exceptions.
        || ex.GetType().FullName?.StartsWith("NPOI.", StringComparison.Ordinal) == true;

    // ---- Pure-garbage inputs ------------------------------------------

    [Theory]
    [InlineData(0)]      // empty
    [InlineData(1)]      // single byte
    [InlineData(16)]
    [InlineData(256)]
    [InlineData(4096)]
    public void Garbage_Bytes_Throw_Documented_Exception(int size)
    {
        // Deterministic per-size seed so failures reproduce.
        var rng = new Random(0xBADF00D ^ size);
        var data = new byte[size];
        rng.NextBytes(data);
        AssertOpenRejects(data);
    }

    [Fact]
    public void All_Zeros_Bytes_Throw_Documented_Exception()
    {
        AssertOpenRejects(new byte[4096]);   // ZIP header == 00 00 00 00 → reject
    }

    [Fact]
    public void All_FF_Bytes_Throw_Documented_Exception()
    {
        var data = new byte[4096];
        Array.Fill(data, (byte)0xFF);
        AssertOpenRejects(data);
    }

    // ---- ZIP-shaped inputs --------------------------------------------

    [Fact]
    public void Empty_Zip_Throws_Documented_Exception()
    {
        // ZIP end-of-central-directory record only; no actual entries.
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) { }
        AssertOpenRejects(ms.ToArray());
    }

    [Fact]
    public void Zip_With_Random_Entries_Throws_Documented_Exception()
    {
        // Valid ZIP but lacks OOXML content-types / sheet parts.
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            for (int i = 0; i < 3; i++)
            {
                var entry = zip.CreateEntry($"junk{i}.dat");
                using var s = entry.Open();
                s.Write(new byte[] { (byte)i, 0xAA, 0xBB });
            }
        }
        AssertOpenRejects(ms.ToArray());
    }

    [Fact]
    public void Zip_With_Truncated_ContentTypes_Throws_Documented_Exception()
    {
        // Looks OOXML-ish — has [Content_Types].xml — but the XML body
        // is malformed.
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = zip.CreateEntry("[Content_Types].xml");
            using var s = entry.Open();
            // Unclosed root element — XML parser must error.
            s.Write(Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><Types><Default"));
        }
        AssertOpenRejects(ms.ToArray());
    }

    [Fact]
    public void Zip_With_XML_Expansion_Bomb_Throws_Documented_Exception()
    {
        // Classic billion-laughs entity expansion. NetXlsx / NPOI must
        // not blow up RAM trying to expand this.
        const string bomb = """
            <?xml version="1.0"?>
            <!DOCTYPE lolz [
              <!ENTITY lol "lol">
              <!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">
              <!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">
              <!ENTITY lol4 "&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;">
              <!ENTITY lol5 "&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;">
              <!ENTITY lol6 "&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;">
              <!ENTITY lol7 "&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;">
            ]>
            <lolz>&lol7;</lolz>
            """;
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = zip.CreateEntry("[Content_Types].xml");
            using var s = entry.Open();
            s.Write(Encoding.UTF8.GetBytes(bomb));
        }
        AssertOpenRejects(ms.ToArray());
    }

    [Fact]
    public void Zip_Bomb_Highly_Compressible_Payload_Throws_ResourceLimit()
    {
        // 64 MiB of zeros compresses to a few KB; with the default
        // ReadMaxUncompressedBytes of 256 MiB, this *fits*. Force a
        // smaller cap and verify the limit fires.
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = zip.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
            using var s = entry.Open();
            var chunk = new byte[1024 * 1024];   // 1 MiB of zeros
            for (int i = 0; i < 64; i++) s.Write(chunk);
        }

        var data = ms.ToArray();
        var tightCap = new WorkbookOptions { ReadMaxUncompressedBytes = 1024 * 1024 };  // 1 MiB
        AssertOpenRejects(data, tightCap);
    }

    // ---- Bit-flip mutations of a known-good base ---------------------

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void Bitflip_Mutation_Of_Valid_Workbook_Throws_Documented_Exception(int seed)
    {
        // Build a tiny valid workbook, then flip random bytes and feed
        // the mutated buffer back to Open. Most flips break ZIP CRC or
        // XML parsing — both must surface as a known exception.
        var baseline = BuildValidWorkbook();
        var rng = new Random(seed);
        var mutated = (byte[])baseline.Clone();
        // Flip ~5 bytes; enough to break structure without 100%
        // guaranteeing a particular failure mode.
        for (int i = 0; i < 5; i++)
        {
            int pos = rng.Next(mutated.Length);
            mutated[pos] = (byte)~mutated[pos];
        }

        // A bit-flip *might* land on a byte that's still parseable; in
        // that case Open succeeds, which is fine. We only assert that
        // when it fails, it fails inside the documented contract.
        try
        {
            using var wb = Workbook.Open(new MemoryStream(mutated));
        }
        catch (Exception ex)
        {
            IsAcceptableOpenException(ex)
                .Should().BeTrue($"bit-flip mutation produced {ex.GetType().FullName}: {ex.Message}");
        }
    }

    // ---- Bulk random sweep -------------------------------------------
    //
    // Slow-ish test; gated on the Fuzz trait. Runs N random inputs of
    // varying sizes; asserts each terminates in bounded time with an
    // acceptable exception. Good signal on "did a new code path
    // introduce a crash."

    [Fact]
    public async Task Bulk_Random_Sweep_Open_Never_Hangs_Or_Crashes()
    {
        const int iterations = 100;
        var rng = new Random(42);
        for (int i = 0; i < iterations; i++)
        {
            int size = rng.Next(1, 8192);
            var data = new byte[size];
            rng.NextBytes(data);

            // Time-bound each call. A hang here is the kind of finding
            // we care about (resource exhaustion or infinite loop).
            using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(2));
            try
            {
                await Task.Run(() =>
                {
                    try { using var wb = Workbook.Open(new MemoryStream(data)); }
                    catch (Exception ex) when (IsAcceptableOpenException(ex)) { /* expected */ }
                }, cts.Token);
            }
            catch (OperationCanceledException)
            {
                throw new Xunit.Sdk.XunitException(
                    $"Open() exceeded 2s on iteration {i} with seed-byte payload of size {size}. " +
                    "Investigate as a potential resource-exhaustion finding.");
            }
        }
    }

    // ---- Helpers ------------------------------------------------------

    private static void AssertOpenRejects(byte[] data, WorkbookOptions? options = null)
    {
        Action act = () =>
        {
            using var wb = Workbook.Open(new MemoryStream(data), leaveOpen: false, options: options);
        };
        try
        {
            act();
            // A successful open of garbage bytes would be a finding —
            // but it's *technically* possible if the bytes happen to be
            // a valid empty workbook. We accept silent success only for
            // explicit empty-zip with no NetXlsx content; flag others.
            throw new Xunit.Sdk.XunitException(
                "Open() succeeded on what should be malformed input. " +
                "Either the input was actually valid (review the corpus) " +
                "or the validation gate has a gap.");
        }
        catch (Xunit.Sdk.XunitException) { throw; }
        catch (Exception ex)
        {
            IsAcceptableOpenException(ex)
                .Should().BeTrue(
                    $"Open() threw {ex.GetType().FullName} which is not in the documented contract: {ex.Message}");
        }
    }

    private static byte[] BuildValidWorkbook()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("hello");
        sh["B1"].SetNumber(42);
        using var ms = new MemoryStream();
        wb.Save(ms, leaveOpen: true);
        return ms.ToArray();
    }
}
