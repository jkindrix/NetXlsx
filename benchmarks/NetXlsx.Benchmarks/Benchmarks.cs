// CI-friendly micro/meso benchmarks that exercise the perf-claims
// recorded in docs/design.md §5. Sized to complete in under a few
// minutes on a typical CI runner so the regression gate can run on
// every PR.

using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Jobs;

namespace NetXlsx.Benchmarks;

/// <summary>
/// Configuration: short job (3 warmup × 3 iterations) so each
/// benchmark completes in ~5-10 seconds, full suite in under a few
/// minutes. CI regression-gate trades statistical precision for
/// throughput; for full statistical runs use the default config.
/// </summary>
public class CiConfig : ManualConfig
{
    public CiConfig()
    {
        AddJob(Job.ShortRun);
        AddExporter(JsonExporter.Brief);
        AddDiagnoser(MemoryDiagnoser.Default);
    }
}

[Config(typeof(CiConfig))]
public class WriteBenchmarks
{
    private readonly string _tmpPath;

    public WriteBenchmarks()
    {
        _tmpPath = Path.Combine(Path.GetTempPath(), $"netxlsx-bench-{Guid.NewGuid():N}.xlsx");
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_tmpPath)) File.Delete(_tmpPath);
    }

    /// <summary>
    /// Cold create + save of an empty workbook. Design §5 target: &lt; 50ms.
    /// Primarily exercises OPC packaging + XSSFWorkbook bootstrap cost.
    /// </summary>
    [Benchmark]
    public void ColdCreateAndSave()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Save(_tmpPath);
    }

    /// <summary>
    /// Write 5,000 rows × 10 cols with mixed types (string + double + int).
    /// Spike 2 measured the in-memory budget threshold at ~30k rows;
    /// 5k is a representative CI-sized sample of the in-memory write path.
    /// </summary>
    [Benchmark]
    public void Write5kRows()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        for (int r = 1; r <= 5_000; r++)
        {
            var row = sheet.AppendRow();
            row.Set(1, $"row-{r}").Set(2, r).Set(3, r * 1.5);
            for (int c = 4; c <= 10; c++) row.Set(c, r + c);
        }
        wb.Save(_tmpPath);
    }

    /// <summary>
    /// Write 10,000 cells with a small style palette (5 distinct styles).
    /// Exercises the CellStylePool dedup path (decision #4 / spike 1).
    /// Design §5 throughput target: &gt; 500k styled cells/s for small
    /// palettes — for 10k cells that's &lt; 20ms.
    /// </summary>
    [Benchmark]
    public void StyledWrite_SmallPalette()
    {
        _ = _tmpPath;  // ensures the method is bound to instance state for CA1822
        var styles = new[]
        {
            new CellStyle { Bold = true },
            new CellStyle { Italic = true },
            new CellStyle { NumberFormat = NumberFormats.Currency },
            new CellStyle { Background = Color.Yellow },
            new CellStyle { Background = Color.LightGray },
        };
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        for (int i = 0; i < 10_000; i++)
        {
            var cell = sheet[i / 100 + 1, i % 100 + 1];
            cell.SetNumber(i);
            cell.Style(styles[i % styles.Length]);
        }
    }

    /// <summary>
    /// Stream 50,000 rows × 10 cols via the SXSSF streaming path.
    /// Spike 2 measured a flat ~70 MB ΔGC at 500k rows; 50k is a
    /// scaled-down CI-friendly variant.
    /// </summary>
    [Benchmark]
    public void StreamingWrite_50kRows()
    {
        using var wb = Workbook.CreateStreaming(new StreamingOptions { RowAccessWindowSize = 1_000 });
        var sheet = wb.AddSheet("Big");
        for (int r = 1; r <= 50_000; r++)
        {
            var row = sheet.AppendRow();
            row.Set(1, r).Set(2, $"row-{r}");
            for (int c = 3; c <= 10; c++) row.Set(c, r + c);
        }
        wb.Save(_tmpPath);
    }
}

[Config(typeof(CiConfig))]
public class ReadBenchmarks
{
    private string _smallFilePath = "";

    [GlobalSetup]
    public void Setup()
    {
        _smallFilePath = Path.Combine(Path.GetTempPath(), $"netxlsx-bench-readsrc-{Guid.NewGuid():N}.xlsx");
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        for (int r = 1; r <= 1_000; r++)
        {
            var row = sheet.AppendRow();
            row.Set(1, $"row-{r}").Set(2, r).Set(3, r * 1.5);
        }
        wb.Save(_smallFilePath);
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_smallFilePath)) File.Delete(_smallFilePath);
    }

    /// <summary>
    /// Open a 1,000-row workbook, sum one column, close. CI-friendly
    /// proxy for the design §5 "open + read 100k × 20" target — the
    /// open cost is dominated by OPC unpack which scales sub-linearly.
    /// </summary>
    [Benchmark]
    public double OpenAndReadColumnSum()
    {
        using var wb = Workbook.Open(_smallFilePath);
        var sheet = wb["Data"];
        double sum = 0;
        for (int r = 1; r <= 1_000; r++)
        {
            sum += sheet[r, 2].GetNumber() ?? 0;
        }
        return sum;
    }
}
