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
        || ex is NotSupportedException;      // unsupported OOXML feature

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
        // Classic billion-laughs entity expansion. The engine must not
        // blow up RAM trying to expand this.
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
        // PR smoke runs the default; the nightly deep-fuzz workflow (R-26)
        // scales this up via NETXLSX_FUZZ_ITERATIONS.
        int iterations = EnvInt("NETXLSX_FUZZ_ITERATIONS", 100);
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

    // ---- Deep bit-flip sweep across the corpus (R-26) ------------------
    //
    // The corpus is COMMITTED AS CODE: deterministic builders with fixed
    // content (below), not binary fixtures — same reproducibility, none of
    // the fixture-provenance overhead, and the corpus automatically tracks
    // the writer (a new emission path joins the corpus by construction
    // when a builder grows). The nightly workflow scales the seed count
    // via NETXLSX_FUZZ_BITFLIP_SEEDS; the PR-smoke default stays small.

    [Fact]
    public async Task Deep_Bitflip_Sweep_Across_Corpus()
    {
        int seeds = EnvInt("NETXLSX_FUZZ_BITFLIP_SEEDS", 5);
        var corpus = BuildCorpus();

        for (int seed = 0; seed < seeds; seed++)
        {
            foreach (var (name, baseline) in corpus)
            {
                var rng = new Random(seed * 31 + name.Length);
                var mutated = (byte[])baseline.Clone();
                int flips = 1 + rng.Next(8);
                for (int i = 0; i < flips; i++)
                {
                    int pos = rng.Next(mutated.Length);
                    mutated[pos] = (byte)~mutated[pos];
                }

                using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(2));
                try
                {
                    await Task.Run(() =>
                    {
                        try { using var wb = Workbook.Open(new MemoryStream(mutated)); }
                        catch (Exception ex) when (IsAcceptableOpenException(ex)) { /* contract */ }
                        catch (Exception ex)
                        {
                            // Corpus bytes are content-deterministic but not
                            // byte-deterministic across runs (zip timestamps),
                            // so a seed alone cannot reproduce a finding —
                            // persist the exact failing input for triage (the
                            // nightly workflow uploads it as an artifact).
                            Directory.CreateDirectory("fuzz-findings");
                            var dump = Path.Combine("fuzz-findings", $"{name}-seed{seed}.xlsx.bin");
                            File.WriteAllBytes(dump, mutated);
                            throw new Xunit.Sdk.XunitException(
                                $"corpus '{name}' seed {seed} ({flips} flips) produced " +
                                $"{ex.GetType().FullName} outside the documented contract: {ex.Message} " +
                                $"[failing input saved to {Path.GetFullPath(dump)}]");
                        }
                    }, cts.Token);
                }
                catch (OperationCanceledException)
                {
                    throw new Xunit.Sdk.XunitException(
                        $"Open() exceeded 2s on corpus '{name}' seed {seed} — potential resource-exhaustion finding.");
                }
            }
        }
    }

    // Deterministic corpus: one workbook per major emission path, fixed
    // content. Built once per test run.
    private static (string Name, byte[] Bytes)[] BuildCorpus()
    {
        static byte[] Build(Action<IWorkbook> fill, Func<IWorkbook>? factory = null)
        {
            using var wb = factory is null ? Workbook.Create() : factory();
            fill(wb);
            using var ms = new MemoryStream();
            wb.Save(ms, leaveOpen: true);
            return ms.ToArray();
        }

        var plain = Build(wb =>
        {
            var s = wb.AddSheet("S");
            s["A1"].SetString("hello");
            s["B2"].SetNumber(42.5);
            s["C3"].SetBool(true);
        });

        var formulaHeavy = Build(wb =>
        {
            var s = wb.AddSheet("F");
            for (int r = 1; r <= 20; r++)
            {
                s[r, 1].SetNumber(r);
                s[r, 2].SetFormula($"=A{r}*2+SUM($A$1:$A${r})");
            }
            wb.AddNamedRange("All", "F!$A$1:$A$20");
        });

        var styledRich = Build(wb =>
        {
            var s = wb.AddSheet("R");
            s["A1"].SetRichText(new RichText(
                new RichTextRun("hot") { Style = new RichTextStyle { Bold = true } },
                new RichTextRun("cold") { Style = new RichTextStyle { Italic = true } }));
            s["B1"].SetString("merged"); s.MergeCells("B1:C2");
            s["D1"].SetDate(new DateTime(2026, 6, 12, 8, 0, 0));
            s["E1"].Comment("note", "fuzz");
            s["F1"].Hyperlink("https://example.com");
        });

        var macro = Build(wb => wb.AddSheet("M")["A1"].SetString("xlsm"),
            () => Workbook.CreateMacroEnabled());

        byte[] streaming;
        using (var ms = new MemoryStream())
        {
            using (var swb = Workbook.CreateStreaming())
            {
                var sh = swb.AddSheet("Big");
                for (int i = 1; i <= 50; i++)
                {
                    var row = sh.AppendRow();
                    row.Set(1, i);
                    row.Set(2, $"row-{i}");
                }
                swb.Save(ms, leaveOpen: true);
            }
            streaming = ms.ToArray();
        }

        return new[]
        {
            ("plain", plain),
            ("formula-heavy", formulaHeavy),
            ("styled-rich", styledRich),
            ("macro-enabled", macro),
            ("streaming", streaming),
        };
    }

    private static int EnvInt(string name, int @default)
        => int.TryParse(Environment.GetEnvironmentVariable(name), out int v) && v > 0 ? v : @default;

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
