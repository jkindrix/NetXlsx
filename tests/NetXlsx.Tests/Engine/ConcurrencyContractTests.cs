// I-82 engine swap — concurrency contract parity (pre-cutover parity slice).
//
// The O-15 scout found the SDK engine missing decision #43's opportunistic
// reentry counter and decision I-59's strict real-lock mode: a raced AddSheet
// corrupted the package and the InvalidOperationException escaped at Dispose.
// These tests pin the SDK engine to the NPOI engine's contract (mirrors
// StrictConcurrencyTests + the #43 default-mode coverage) so the cutover
// inherits it green.

using System;
using System.Threading;
using System.Threading.Tasks;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class ConcurrencyContractTests
{
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
    public void Strict_Mode_Permits_Sequential_Mixed_Mutations()
    {
        using var wb = Workbook.Create(new WorkbookOptions { StrictConcurrencyDetection = true });
        wb.AddSheet("Outer");
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
        // Default mode throws on concurrent mutation; strict mode serializes
        // them via the real lock. 50 parallel AddSheet calls must all complete
        // and produce 50 sheets — on the SDK engine this also proves the part
        // graph is not corrupted by racing AddNewPart calls.
        using var wb = Workbook.Create(new WorkbookOptions { StrictConcurrencyDetection = true });
        const int n = 50;
        using var barrier = new Barrier(n);
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
        // The opportunistic counter (#43): a racing AddSheet that loses throws
        // before mutating, so the workbook ends in a consistent state.
        using var wb = Workbook.Create();   // default: StrictConcurrencyDetection = false
        var saw = 0;
        using var barrier = new Barrier(2);
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
        // A contention hit isn't guaranteed every run; what we must not see is
        // silent corruption — the workbook ends consistent either way.
        wb.SheetCount.Should().Be(2 - saw);
    }

    [Fact]
    public async Task Default_Mode_Concurrent_Mutation_Throws_With_Discoverable_Message()
    {
        // When the counter trips, the message must point users at the
        // strict-mode option (decision I-59 discoverability) and state that
        // workbooks aren't thread-safe.
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
}
