// Extended benchmark coverage per decision I-62: micro (per-cell),
// macro (multi-sheet / large workbook), and percentile reporting.
// Closes the v1.0 external-review recommendation #2 ("extend
// benchmark suite to micro + macro shapes with percentiles").
//
// All extended benchmarks use CiConfig so the regression-gate
// captures them on every PR alongside the original meso suite.

using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Jobs;

namespace NetXlsx.Benchmarks;

/// <summary>
/// Like <see cref="CiConfig"/> but adds percentile-column output
/// (P50 / P95 / P99). Used for the extended micro+macro suite so
/// long-tail latency is visible in the regression gate, not just
/// mean / median.
/// </summary>
public class CiConfigWithPercentiles : ManualConfig
{
    public CiConfigWithPercentiles()
    {
        AddJob(Job.ShortRun);
        AddExporter(JsonExporter.Brief);
        AddDiagnoser(MemoryDiagnoser.Default);
        // BDN adds Mean/Error/StdDev by default. The brief JSON
        // exporter records full Statistics — P50/P95/P99 are part
        // of the JSON payload and consumed by the regression-gate
        // script (bench-compare).
    }
}

/// <summary>
/// Micro benchmarks — per-cell write cost for each scalar type.
/// Establishes a baseline for the cheapest possible operation so
/// regressions in the hot path show up immediately rather than
/// being lost in the noise of larger workloads.
/// </summary>
[Config(typeof(CiConfigWithPercentiles))]
public class MicroBenchmarks
{
    private IWorkbook _wb = null!;
    private ISheet _sheet = null!;

    [IterationSetup]
    public void Setup()
    {
        _wb = Workbook.Create();
        _sheet = _wb.AddSheet("S");
    }

    [IterationCleanup]
    public void Cleanup()
    {
        _wb.Dispose();
    }

    /// <summary>Single string write, no styling, no save.</summary>
    [Benchmark]
    public void Cell_SetString()
    {
        _sheet["A1"].SetString("hello");
    }

    /// <summary>Single number write.</summary>
    [Benchmark]
    public void Cell_SetNumber_Double()
    {
        _sheet["A1"].SetNumber(3.14);
    }

    /// <summary>Single int write through the int overload.</summary>
    [Benchmark]
    public void Cell_SetNumber_Int()
    {
        _sheet["A1"].SetNumber(42);
    }

    /// <summary>Single bool write.</summary>
    [Benchmark]
    public void Cell_SetBool()
    {
        _sheet["A1"].SetBool(true);
    }

    /// <summary>Single DateTime write — includes default-style application.</summary>
    [Benchmark]
    public void Cell_SetDate()
    {
        _sheet["A1"].SetDate(new DateTime(2026, 5, 22));
    }

    /// <summary>Style application of a fresh CellStyle (forces a pool miss).</summary>
    [Benchmark]
    public void Cell_Style_FreshCellStyle()
    {
        // Each iteration uses a distinct CellStyle so the pool can't dedup.
        // Worst-case for style allocation.
        _sheet["A1"].Style(new CellStyle { Bold = true, FontSize = (DateTime.UtcNow.Millisecond % 100) + 1 });
    }

    /// <summary>Style application of a pool-hit CellStyle.</summary>
    [Benchmark]
    public void Cell_Style_PoolHit()
    {
        // Same CellStyle every iteration; second iteration onward hits.
        _sheet["A1"].Style(new CellStyle { Bold = true });
    }

    /// <summary>A1-address parse — runs on every <c>sheet["A1"]</c> indexer call.</summary>
    [Benchmark]
    public (int, int) CellAddress_ParseA1()
    {
        // Returning the result so CA1822 doesn't flag the method as static-able
        // — and so the JIT can't trivially elide the call.
        _ = _sheet;
        return CellAddress.Parse("AB123");
    }

    /// <summary>1-based row/col → A1 format (the inverse).</summary>
    [Benchmark]
    public string CellAddress_FormatA1()
    {
        _ = _sheet;
        return CellAddress.Format(123, 28);
    }
}

/// <summary>
/// Macro benchmarks — workbook shapes that surface scaling issues
/// invisible in the meso (5k-row) tier. Specifically: many-sheets
/// and very-wide layouts.
/// </summary>
[Config(typeof(CiConfigWithPercentiles))]
public class MacroBenchmarks
{
    private string _tmpPath = null!;

    [GlobalSetup]
    public void Setup()
    {
        _tmpPath = Path.Combine(Path.GetTempPath(), $"netxlsx-macro-{Guid.NewGuid():N}.xlsx");
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_tmpPath)) File.Delete(_tmpPath);
    }

    /// <summary>
    /// 500 sheets, one row each. Stresses sheet-creation overhead
    /// and the workbook's internal sheet collection, separate from
    /// per-cell cost.
    /// </summary>
    [Benchmark]
    public void Macro_500Sheets()
    {
        using var wb = Workbook.Create();
        for (int s = 0; s < 500; s++)
        {
            // Sheet name capped at 31 chars; "Sheet" + 0..499 fits.
            var sh = wb.AddSheet($"Sheet{s}");
            sh.AppendRow().Set(1, s);
        }
        wb.Save(_tmpPath);
    }

    /// <summary>
    /// One sheet, 1,000 rows × 100 cols (= 100k cells). Tests
    /// the random-access path at a scale where row-cache effects
    /// matter.
    /// </summary>
    [Benchmark]
    public void Macro_100kCells_Wide()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        for (int r = 1; r <= 1_000; r++)
        {
            var row = sh.AppendRow();
            for (int c = 1; c <= 100; c++) row.Set(c, r * c);
        }
        wb.Save(_tmpPath);
    }

    /// <summary>
    /// Streaming write at 200k rows × 10 cols (= 2M cells), the
    /// scale where streaming starts to matter. Spike 2 measured
    /// memory budget at 500k rows; this is half that, sized for
    /// CI runtime.
    /// </summary>
    [Benchmark]
    public void Macro_Streaming_200kRows()
    {
        using var wb = Workbook.CreateStreaming(new StreamingOptions { RowAccessWindowSize = 1_000 });
        var sh = wb.AddSheet("Big");
        for (int r = 1; r <= 200_000; r++)
        {
            var row = sh.AppendRow();
            row.Set(1, r).Set(2, $"r{r}");
            for (int c = 3; c <= 10; c++) row.Set(c, r + c);
        }
        wb.Save(_tmpPath);
    }
}

/// <summary>
/// Read-side micro benchmark — round-trip a 1×1 workbook through
/// Save → Open → GetString. Establishes the floor cost of opening
/// a workbook, distinct from the cost of reading lots of data.
/// </summary>
[Config(typeof(CiConfigWithPercentiles))]
public class ReadMicroBenchmarks
{
    private byte[] _smallWb = null!;

    [GlobalSetup]
    public void Setup()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("hello");
        using var ms = new MemoryStream();
        wb.Save(ms, leaveOpen: true);
        _smallWb = ms.ToArray();
    }

    /// <summary>Open a 1-cell workbook from memory, read the value, dispose.</summary>
    [Benchmark]
    public string Open_OneCell_Read()
    {
        using var wb = Workbook.Open(new MemoryStream(_smallWb));
        return wb[0]["A1"].GetString();
    }
}
