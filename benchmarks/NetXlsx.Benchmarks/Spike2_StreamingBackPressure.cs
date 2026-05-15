// Spike 2 — Streaming back-pressure
//
// Question: at what row count does the in-memory XSSFWorkbook peak working
// set exceed a threshold that would push streaming write (SXSSF) into v1
// instead of v2? Specifically: peak working set for {10k, 50k, 100k, 250k,
// 500k} rows × 20 columns of mixed cell types.
//
// Resolves: roadmap pre-impl spike "Streaming back-pressure". Sets the
// rough threshold at which callers should reach for v2's streaming writer.
//
// Run: dotnet run --project benchmarks/NetXlsx.Benchmarks -c Release -- spike-2

using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;

namespace NetXlsx.Benchmarks;

internal static class Spike2_StreamingBackPressure
{
    public static int Run()
    {
        Console.WriteLine("Spike 2 — streaming back-pressure (NPOI 2.7.3, .NET 9)");
        Console.WriteLine($"  Process model: peak WorkingSet64 sampled at 50ms intervals during write+save");
        Console.WriteLine();

        int[] rowCounts = { 10_000, 50_000, 100_000, 250_000 };
        int cols = 20;

        Console.WriteLine($"{"Mode",-12} {"Rows",10} {"Cols",6} {"Time (s)",12} {"ΔWS (MB)",14} {"ΔGC (MB)",14} {"File (MB)",12}");
        Console.WriteLine(new string('-', 86));

        foreach (var rows in rowCounts)
        {
            MeasureXssfInMemory(rows, cols);
        }

        // SXSSF streaming write — same row counts plus a 500k stress
        int[] sxssfRowCounts = { 10_000, 50_000, 100_000, 250_000, 500_000 };
        foreach (var rows in sxssfRowCounts)
        {
            MeasureSxssfStreaming(rows, cols);
        }

        return 0;
    }

    private static void MeasureXssfInMemory(int rows, int cols)
    {
        using var sampler = new PeakWorkingSetSampler();
        var sw = Stopwatch.StartNew();
        var path = Path.Combine(Path.GetTempPath(), $"spike2-xssf-{rows}.xlsx");
        try
        {
            using var wb = new XSSFWorkbook();
            var sheet = wb.CreateSheet("Data");
            for (int r = 0; r < rows; r++)
            {
                var row = sheet.CreateRow(r);
                for (int c = 0; c < cols; c++)
                {
                    var cell = row.CreateCell(c);
                    // Mixed types: alternate string / number / number
                    if (c % 3 == 0) cell.SetCellValue($"r{r}c{c}");
                    else cell.SetCellValue(r * cols + c);
                }
            }
            using var fs = File.Create(path);
            wb.Write(fs);
            sw.Stop();
            sampler.Stop();
            var sizeMb = new FileInfo(path).Length / 1024.0 / 1024.0;
            Console.WriteLine($"{"XSSF",-12} {rows,10:N0} {cols,6} {sw.Elapsed.TotalSeconds,12:F2} {sampler.PeakMb,14:F1} {sampler.PeakManagedMb,14:F1} {sizeMb,12:F2}");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static void MeasureSxssfStreaming(int rows, int cols)
    {
        using var sampler = new PeakWorkingSetSampler();
        var sw = Stopwatch.StartNew();
        var path = Path.Combine(Path.GetTempPath(), $"spike2-sxssf-{rows}.xlsx");
        try
        {
            // Row window of 100 — beyond that, NPOI flushes rows to disk.
            using var wb = new SXSSFWorkbook(100);
            var sheet = wb.CreateSheet("Data");
            for (int r = 0; r < rows; r++)
            {
                var row = sheet.CreateRow(r);
                for (int c = 0; c < cols; c++)
                {
                    var cell = row.CreateCell(c);
                    if (c % 3 == 0) cell.SetCellValue($"r{r}c{c}");
                    else cell.SetCellValue(r * cols + c);
                }
            }
            using var fs = File.Create(path);
            wb.Write(fs);
            wb.Dispose();
            sw.Stop();
            sampler.Stop();
            var sizeMb = new FileInfo(path).Length / 1024.0 / 1024.0;
            Console.WriteLine($"{"SXSSF",-12} {rows,10:N0} {cols,6} {sw.Elapsed.TotalSeconds,12:F2} {sampler.PeakMb,14:F1} {sampler.PeakManagedMb,14:F1} {sizeMb,12:F2}");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private sealed class PeakWorkingSetSampler : IDisposable
    {
        private readonly CancellationTokenSource _cts = new();
        private readonly Task _task;
        private readonly long _baselineBytes;
        private long _peakBytes;
        // Managed-memory peak (GC.GetTotalMemory) — measures the actual
        // managed-heap working set, unaffected by OS WS retention.
        private readonly long _baselineGc;
        private long _peakGc;

        public double PeakMb => Math.Max(0, _peakBytes - _baselineBytes) / 1024.0 / 1024.0;
        public double PeakManagedMb => Math.Max(0, _peakGc - _baselineGc) / 1024.0 / 1024.0;

        public PeakWorkingSetSampler()
        {
            // Force GC so the baseline reflects only live state.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            _baselineBytes = Process.GetCurrentProcess().WorkingSet64;
            _peakBytes = _baselineBytes;
            _baselineGc = GC.GetTotalMemory(false);
            _peakGc = _baselineGc;

            _task = Task.Run(async () =>
            {
                try
                {
                    while (!_cts.Token.IsCancellationRequested)
                    {
                        await Task.Delay(50, _cts.Token).ConfigureAwait(false);
                        var ws = Process.GetCurrentProcess().WorkingSet64;
                        if (ws > _peakBytes) _peakBytes = ws;
                        var gc = GC.GetTotalMemory(false);
                        if (gc > _peakGc) _peakGc = gc;
                    }
                }
                catch (TaskCanceledException) { }
            });
        }

        public void Stop()
        {
            _cts.Cancel();
            try { _task.Wait(500); } catch { /* swallow */ }
        }

        public void Dispose() => Stop();
    }
}
