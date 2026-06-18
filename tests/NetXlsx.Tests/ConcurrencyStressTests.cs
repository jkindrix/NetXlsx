// R-30 remainder — thread-safety stress beyond the detection contract.
//
// StrictConcurrencyTests / ConcurrencyContractTests already prove the
// *contract*: default mode surfaces the concurrency guard as
// InvalidOperationException, strict mode serializes concurrent mutations
// without throwing, and dispose-vs-mutate never tears state. What those don't
// assert is that the workbook produced by heavy *mixed* concurrent mutation is
// actually a valid, complete, persistable file. This slice closes that gap:
// hammer one strict-mode workbook from many threads with a mix of operations,
// then save → reopen and verify every mutation is present and intact.

using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class ConcurrencyStressTests
{
    [Fact]
    public async Task Strict_Mode_Mixed_Mutation_Stress_Roundtrips_Intact()
    {
        const int threads = 8;
        const int perThread = 25;   // 200 sheets total

        using var wb = Workbook.Create(new WorkbookOptions { StrictConcurrencyDetection = true });
        using var barrier = new Barrier(threads);
        var failures = new ConcurrentQueue<Exception>();

        var tasks = Enumerable.Range(0, threads).Select(t => Task.Run(() =>
        {
            barrier.SignalAndWait();   // release all threads together for contention
            try
            {
                for (int i = 0; i < perThread; i++)
                {
                    // A mix of mutating paths that each enter the workbook's
                    // mutation scope: add a sheet, write two cells on it, and
                    // register a workbook-scoped named range pointing at it.
                    var sheet = wb.AddSheet($"S{t}_{i}");
                    sheet["A1"].SetString($"t{t}i{i}");
                    sheet["B1"].SetNumber(t * 1000 + i);
                    wb.AddNamedRange($"N{t}_{i}", $"'S{t}_{i}'!$A$1");
                }
            }
            catch (Exception ex) { failures.Enqueue(ex); }
        })).ToArray();

        await Task.WhenAll(tasks);

        failures.Should().BeEmpty("strict mode must serialize every mutation without throwing");
        wb.SheetCount.Should().Be(threads * perThread);
        wb.NamedRanges.Should().HaveCount(threads * perThread);

        // The decisive check: the concurrently-built workbook is a valid,
        // complete file. Save and reopen, then confirm every sheet and its
        // cell payload survived intact.
        using var ms = new System.IO.MemoryStream();
        wb.Save(ms, leaveOpen: true);
        ms.Position = 0;
        using var reopened = Workbook.Open(ms);

        reopened.SheetCount.Should().Be(threads * perThread);
        reopened.NamedRanges.Should().HaveCount(threads * perThread);
        for (int t = 0; t < threads; t++)
        {
            for (int i = 0; i < perThread; i++)
            {
                var sheet = reopened[$"S{t}_{i}"];
                sheet["A1"].GetString().Should().Be($"t{t}i{i}");
                sheet["B1"].GetNumber().Should().Be(t * 1000 + i);
            }
        }
    }

    [Fact]
    public async Task Strict_Mode_Concurrent_Save_And_Mutate_Surfaces_Only_Contract_Exceptions()
    {
        // Save participates in the strict lock (like every mutating path), so a
        // save racing mutations must serialize cleanly: each Save produces a
        // valid package and each mutation either lands or waits — never a torn
        // half-saved document or an off-contract engine exception.
        using var wb = Workbook.Create(new WorkbookOptions { StrictConcurrencyDetection = true });
        wb.AddSheet("Seed")["A1"].SetString("seed");

        var unexpected = new ConcurrentQueue<Exception>();
        using var stop = new CancellationTokenSource(TimeSpan.FromMilliseconds(500));

        var mutator = Task.Run(() =>
        {
            int n = 0;
            while (!stop.IsCancellationRequested)
            {
                try { wb.AddSheet($"M{n++}"); }
                catch (ObjectDisposedException) { return; }
                catch (Exception ex) { unexpected.Enqueue(ex); return; }
            }
        });

        var saver = Task.Run(() =>
        {
            while (!stop.IsCancellationRequested)
            {
                try
                {
                    using var ms = new System.IO.MemoryStream();
                    wb.Save(ms);
                }
                catch (ObjectDisposedException) { return; }
                catch (Exception ex) { unexpected.Enqueue(ex); return; }
            }
        });

        await Task.WhenAll(mutator, saver);

        unexpected.Should().BeEmpty("save-vs-mutate under the strict lock must stay on-contract");

        // The workbook is still coherent after the race: a final save reopens cleanly.
        using var final = new System.IO.MemoryStream();
        wb.Save(final, leaveOpen: true);
        final.Position = 0;
        using var reopened = Workbook.Open(final);
        reopened.TryGetSheet("Seed", out _).Should().BeTrue();
    }
}
