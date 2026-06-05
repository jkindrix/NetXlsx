// Spike 3 — Async wrapping cost
//
// HISTORICAL (2026-06-04): this spike measured *NPOI* (the pre-v2.0.0
// engine) and resolved design decision #5's SaveAsync shape. The v2.0.0
// engine swap (I-82) retired NPOI from the library; the NPOI reference
// below is deliberate — it preserves the spike as measured evidence, not
// part of any shipped closure.
//
// Question: is `Task.Run`-wrapped NPOI save measurably better than sync
// save for representative workbook sizes? Specifically: does it improve
// p95 latency under concurrent load without introducing non-trivial
// thread-pool starvation?
//
// Resolves: design decision #5 — whether `SaveAsync` should `Task.Run`
// internally, pass through to the FileStream async path, or take a hybrid.
//
// Run: dotnet run --project benchmarks/NetXlsx.Benchmarks -c Release -- spike-3

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;

namespace NetXlsx.Benchmarks;

internal static class Spike3_AsyncWrappingCost
{
    public static int Run()
    {
        Console.WriteLine("Spike 3 — async wrapping cost (NPOI 2.7.3, .NET 9)");
        Console.WriteLine();

        int[] rowCounts = { 100, 1_000, 10_000 };
        const int iterations = 20;
        const int concurrentSaves = 10;

        Console.WriteLine("== Single-threaded ==");
        Console.WriteLine($"{"Mode",-18} {"Rows",10} {"Median (ms)",14} {"P95 (ms)",12} {"Mean (ms)",12}");
        Console.WriteLine(new string('-', 70));

        foreach (var rows in rowCounts)
        {
            MeasureSingle(rows, iterations, "Sync", SaveSync);
            MeasureSingle(rows, iterations, "Async (Task.Run)", SaveTaskRun);
        }

        Console.WriteLine();
        Console.WriteLine($"== Concurrent ({concurrentSaves} parallel saves) ==");
        Console.WriteLine($"{"Mode",-18} {"Rows",10} {"P95 (ms)",12} {"Total (ms)",12}");
        Console.WriteLine(new string('-', 60));

        foreach (var rows in rowCounts)
        {
            MeasureConcurrent(rows, concurrentSaves, "Sync", (r, p) => { SaveSync(r, p); return Task.CompletedTask; });
            MeasureConcurrent(rows, concurrentSaves, "Async (Task.Run)", SaveTaskRunAsync);
        }

        return 0;
    }

    private static void MeasureSingle(int rows, int iterations, string label, Action<int, string> save)
    {
        // Warm up
        var warmPath = Path.Combine(Path.GetTempPath(), $"spike3-warm.xlsx");
        save(rows, warmPath);
        if (File.Exists(warmPath)) File.Delete(warmPath);

        var samples = new List<double>(iterations);
        for (int i = 0; i < iterations; i++)
        {
            var path = Path.Combine(Path.GetTempPath(), $"spike3-{label}-{i}.xlsx");
            var sw = Stopwatch.StartNew();
            save(rows, path);
            sw.Stop();
            samples.Add(sw.Elapsed.TotalMilliseconds);
            if (File.Exists(path)) File.Delete(path);
        }

        samples.Sort();
        var median = samples[samples.Count / 2];
        var p95 = samples[(int)(samples.Count * 0.95)];
        var mean = samples.Average();
        Console.WriteLine($"{label,-18} {rows,10:N0} {median,14:F1} {p95,12:F1} {mean,12:F1}");
    }

    private static void MeasureConcurrent(int rows, int parallel, string label, Func<int, string, Task> save)
    {
        // Warm
        Task.WaitAll(Enumerable.Range(0, parallel).Select(i =>
            save(rows, Path.Combine(Path.GetTempPath(), $"spike3-conc-warm-{i}.xlsx"))).ToArray());
        for (int i = 0; i < parallel; i++)
        {
            var w = Path.Combine(Path.GetTempPath(), $"spike3-conc-warm-{i}.xlsx");
            if (File.Exists(w)) File.Delete(w);
        }

        var perTaskTimes = new double[parallel];
        var totalSw = Stopwatch.StartNew();
        var tasks = new Task[parallel];
        for (int i = 0; i < parallel; i++)
        {
            int idx = i;
            tasks[i] = Task.Run(async () =>
            {
                var sw = Stopwatch.StartNew();
                await save(rows, Path.Combine(Path.GetTempPath(), $"spike3-conc-{label}-{idx}.xlsx")).ConfigureAwait(false);
                sw.Stop();
                perTaskTimes[idx] = sw.Elapsed.TotalMilliseconds;
            });
        }
        Task.WaitAll(tasks);
        totalSw.Stop();

        for (int i = 0; i < parallel; i++)
        {
            var p = Path.Combine(Path.GetTempPath(), $"spike3-conc-{label}-{i}.xlsx");
            if (File.Exists(p)) File.Delete(p);
        }

        Array.Sort(perTaskTimes);
        var p95 = perTaskTimes[(int)(perTaskTimes.Length * 0.95)];
        Console.WriteLine($"{label,-18} {rows,10:N0} {p95,12:F1} {totalSw.Elapsed.TotalMilliseconds,12:F1}");
    }

    private static Task SaveTaskRunAsync(int rows, string path) =>
        Task.Run(() => SaveSync(rows, path));

    private static void SaveTaskRun(int rows, string path)
    {
        // Run-and-wait — measuring the offload overhead inside a "single-threaded" caller.
        Task.Run(() => SaveSync(rows, path)).GetAwaiter().GetResult();
    }

    private static void SaveSync(int rows, string path)
    {
        using var wb = new XSSFWorkbook();
        var sheet = wb.CreateSheet("Data");
        for (int r = 0; r < rows; r++)
        {
            var row = sheet.CreateRow(r);
            for (int c = 0; c < 10; c++)
            {
                var cell = row.CreateCell(c);
                if (c % 2 == 0) cell.SetCellValue($"r{r}c{c}");
                else cell.SetCellValue(r * 10 + c);
            }
        }
        using var fs = File.Create(path);
        wb.Write(fs);
    }
}
