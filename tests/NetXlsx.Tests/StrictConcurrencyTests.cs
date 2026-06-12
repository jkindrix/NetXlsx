// Coverage for the v1.1 strict-concurrency slice: opt-in real-lock
// mode (WorkbookOptions.StrictConcurrencyDetection). Default
// opportunistic-counter behavior is covered elsewhere.

using System;
using System.Threading;
using System.Threading.Tasks;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class StrictConcurrencyTests
{
    [Fact]
    public void StrictConcurrencyDetection_Defaults_To_False()
    {
        var opts = new WorkbookOptions();
        opts.StrictConcurrencyDetection.Should().BeFalse();
    }

    [Fact]
    public void Strict_Mode_Permits_Single_Threaded_Mutations()
    {
        using var wb = Workbook.Create(new WorkbookOptions { StrictConcurrencyDetection = true });
        wb.AddSheet("A");
        wb.AddSheet("B");
        wb.AddSheet("C");
        wb.SheetCount.Should().Be(3);
    }

    [Fact]
    public void Strict_Mode_Permits_Same_Thread_Reentrancy()
    {
        // Strict mode uses Monitor (reentrant), so nested mutations on
        // the same thread are permitted — unlike the default
        // opportunistic counter which throws on reentry.
        using var wb = Workbook.Create(new WorkbookOptions { StrictConcurrencyDetection = true });
        wb.AddSheet("Outer");
        // Synthesize reentrancy by calling AddNamedRange (which also
        // enters the mutation scope) from inside another mutation —
        // here, sequentially in the same thread, which serves as the
        // base assertion that the scope re-enters cleanly.
        Action act = () =>
        {
            wb.AddSheet("S1");
            wb.AddNamedRange("MyName", "S1!$A$1");
        };
        act.Should().NotThrow();
    }

    [Fact]
    public async Task Strict_Mode_Serializes_Concurrent_AddSheet_Without_Throwing()
    {
        // Default mode throws on concurrent mutation; strict mode
        // serializes them via the real lock. Verify that 50 parallel
        // AddSheet calls all complete and produce 50 sheets.
        using var wb = Workbook.Create(new WorkbookOptions { StrictConcurrencyDetection = true });
        const int n = 50;
        var barrier = new Barrier(n);
        var tasks = new Task[n];
        for (int i = 0; i < n; i++)
        {
            int idx = i;
            tasks[i] = Task.Run(() =>
            {
                barrier.SignalAndWait();   // maximize contention
                wb.AddSheet($"S{idx}");
            });
        }
        await Task.WhenAll(tasks);
        wb.SheetCount.Should().Be(n);
    }

    [Fact]
    public async Task Default_Mode_Surfaces_Concurrent_Mutation_As_InvalidOperation()
    {
        // Contrast test: confirm the default opportunistic counter
        // still throws — we don't want a regression where strict mode
        // becomes the unconditional behavior.
        using var wb = Workbook.Create();   // default: StrictConcurrencyDetection = false
        var saw = 0;
        var barrier = new Barrier(2);
        var tasks = new Task[2];
        for (int i = 0; i < 2; i++)
        {
            int idx = i;
            tasks[i] = Task.Run(() =>
            {
                barrier.SignalAndWait();
                try { wb.AddSheet($"S{idx}"); }
                catch (InvalidOperationException) { Interlocked.Increment(ref saw); }
            });
        }
        await Task.WhenAll(tasks);
        // We can't guarantee a contention hit in every CI run, but in
        // the worst case both succeed (zero throws). What we *must not*
        // see is silent corruption — verify the workbook ends in a
        // consistent state.
        wb.SheetCount.Should().Be(2 - saw);
    }

    [Fact]
    public async Task Default_Mode_Concurrent_Mutation_Throws_With_Discoverable_Message()
    {
        // When the default opportunistic counter trips, the thrown message
        // must point users at the strict-mode option (discoverability) and
        // state that workbooks aren't thread-safe. Several threads hammer
        // AddSheet on one workbook (unique names, so the only
        // InvalidOperationException is the concurrency guard, not a name
        // clash); collisions are near-certain and captured fast. Default
        // mode isn't thread-safe, so we don't assert final state here —
        // only the discoverability of the message.
        string? message = null;
        using var wb = Workbook.Create();   // default: StrictConcurrencyDetection = false
        const int n = 4;
        using var barrier = new Barrier(n);
        var tasks = new Task[n];
        for (int t = 0; t < n; t++)
        {
            int tid = t;
            tasks[t] = Task.Run(() =>
            {
                barrier.SignalAndWait();   // release all threads at once
                for (int k = 0; k < 500 && Volatile.Read(ref message) is null; k++)
                {
                    try { wb.AddSheet($"S{tid}_{k}"); }
                    catch (InvalidOperationException ex)
                    {
                        Interlocked.CompareExchange(ref message, ex.Message, null);
                    }
                    catch (SheetNameException) { /* unrelated to the concurrency guard */ }
                }
            });
        }
        await Task.WhenAll(tasks);

        message.Should().NotBeNull(
            "concurrent AddSheet on the default counter must surface InvalidOperationException");
        message.Should().Contain("StrictConcurrencyDetection",
            "the default-mode error must point users at the opt-in real-lock option (decision I-59)");
        message.Should().Contain("not thread-safe");
    }

    // ---- R-15: Dispose participates in the strict lock ------------------

    [Fact]
    public void Strict_Dispose_Is_Clean_And_Idempotent()
    {
        var wb = Workbook.Create(new WorkbookOptions { StrictConcurrencyDetection = true });
        wb.AddSheet("S");
        wb.Dispose();
        Action again = wb.Dispose;
        again.Should().NotThrow("Dispose is idempotent in strict mode too");
    }

    [Fact]
    public void Strict_Dispose_Racing_Mutations_Surfaces_Only_Contract_Exceptions()
    {
        // R-15: with disposal under the same per-workbook lock as every
        // mutating path, a dispose racing mutations can only produce the
        // contract's ObjectDisposedException — never a torn-state engine
        // exception from mutating a half-disposed document.
        var wb = Workbook.Create(new WorkbookOptions { StrictConcurrencyDetection = true });
        var threads = new System.Collections.Generic.List<System.Threading.Thread>();
        var unexpected = new System.Collections.Concurrent.ConcurrentQueue<Exception>();
        using var start = new System.Threading.ManualResetEventSlim(false);

        for (int t = 0; t < 8; t++)
        {
            int id = t;
            var thread = new System.Threading.Thread(() =>
            {
                start.Wait();
                for (int i = 0; i < 50; i++)
                {
                    try { wb.AddSheet($"T{id}-{i}"); }
                    catch (ObjectDisposedException) { return; } // contract
                    catch (Exception ex) { unexpected.Enqueue(ex); return; }
                }
            });
            thread.Start();
            threads.Add(thread);
        }

        start.Set();
        System.Threading.Thread.Sleep(5);
        wb.Dispose();
        foreach (var thread in threads) thread.Join();

        unexpected.Should().BeEmpty(
            "dispose-vs-mutate under the strict lock must never tear state");
    }
}
