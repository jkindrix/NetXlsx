// Comparative benchmarks (ledger R-29). These measure NetXlsx against the
// two 2026 reference points named in the ledger:
//
//   * SpreadCheetah — the closest analog (streaming / source-gen / AOT / MIT).
//     Write-only and streaming-only, so it is compared ONLY on the
//     streaming-write path (its README's 100k×10 ≈ 33 ms reference point).
//     Pitting it on read or buffered DOM features it does not have would be
//     dishonest, so we don't.
//   * ClosedXML — the breadth comparator (full DOM read + write). Compared on
//     the buffered-write path and the read path.
//
// THESE CLASSES ARE DELIBERATELY OUTSIDE THE `NetXlsx.Benchmarks` NAMESPACE.
// The CI regression gate (bench.yml) selects classes with
// `--filter "*Benchmarks*"`, and BenchmarkDotNet's filter matches the FULL
// benchmark id — namespace included. Living in `NetXlsx.Comparative` keeps a
// competitor's run-to-run variance from ever reddening our own absolute-number
// gate (decision S3 / I-87). They run only on an explicit filter, e.g.:
//
//   dotnet run -c Release --framework net10.0 -- --filter '*Comparative*'
//
// Enclosing-namespace lookup still resolves the NetXlsx public types
// (Workbook, StreamingOptions, …) without a `using` — the same mechanism the
// in-namespace benchmarks rely on (ImplicitUsings is disabled project-wide).

using System;
using System.IO;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Order;
using ClosedXML.Excel;
using SpreadCheetah;

namespace NetXlsx.Comparative;

/// <summary>
/// Shared config for the comparative suite. ShortRun (3×3) keeps a full
/// cross-engine sweep under a few minutes even with ClosedXML's heavy DOM in
/// the mix; <see cref="MemoryDiagnoser"/> captures the allocation story
/// (NetXlsx/SpreadCheetah streaming vs ClosedXML's whole-workbook DOM) that is
/// half the point of the comparison. The GitHub-markdown exporter emits a
/// paste-ready table for docs; the brief JSON is for ad-hoc inspection — these
/// results are never fed to the CI baseline (see the file header).
/// </summary>
public class ComparativeConfig : ManualConfig
{
    public ComparativeConfig()
    {
        AddJob(Job.ShortRun);
        AddDiagnoser(MemoryDiagnoser.Default);
        AddExporter(MarkdownExporter.GitHub);
        AddExporter(JsonExporter.Brief);
        // Sort by the declared baseline so the NetXlsx row leads each table and
        // the Ratio column reads "competitor relative to NetXlsx".
        WithOrderer(new DefaultOrderer(SummaryOrderPolicy.Declared));
    }
}

/// <summary>
/// The headline streaming-write comparison: 100,000 rows × 10 columns of
/// mixed scalars, streamed to a temp file. This is SpreadCheetah's home turf
/// and the README reference point (≈33 ms on their hardware). NetXlsx's
/// streaming entry point (<c>CreateStreaming</c>) is the apples-to-apples
/// path; ClosedXML is excluded here — it has no streaming model.
/// </summary>
[Config(typeof(ComparativeConfig))]
public class ComparativeStreamingWrite
{
    private const int Rows = 100_000;
    private const int Cols = 10;

    private string _tmpPath = null!;

    [GlobalSetup]
    public void Setup() =>
        _tmpPath = Path.Combine(Path.GetTempPath(), $"netxlsx-cmp-stream-{Guid.NewGuid():N}.xlsx");

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_tmpPath)) File.Delete(_tmpPath);
    }

    /// <summary>NetXlsx streaming write — the baseline every Ratio is measured against.</summary>
    [Benchmark(Baseline = true, Description = "NetXlsx (streaming)")]
    public void NetXlsx_Streaming()
    {
        using var wb = Workbook.CreateStreaming(new StreamingOptions { RowAccessWindowSize = 1_000 });
        var sheet = wb.AddSheet("Data");
        for (int r = 1; r <= Rows; r++)
        {
            var row = sheet.AppendRow();
            row.Set(1, $"row-{r}").Set(2, r).Set(3, r * 1.5);
            for (int c = 4; c <= Cols; c++) row.Set(c, r + c);
        }
        wb.Save(_tmpPath);
    }

    /// <summary>SpreadCheetah streaming write of the identical workload.</summary>
    [Benchmark(Description = "SpreadCheetah")]
    public async Task SpreadCheetah_Streaming()
    {
        await using var stream = File.Create(_tmpPath);
        await using var sheet = await Spreadsheet.CreateNewAsync(stream);
        await sheet.StartWorksheetAsync("Data");
        var cells = new DataCell[Cols];
        for (int r = 1; r <= Rows; r++)
        {
            cells[0] = new DataCell($"row-{r}");
            cells[1] = new DataCell(r);
            cells[2] = new DataCell(r * 1.5);
            for (int c = 4; c <= Cols; c++) cells[c - 1] = new DataCell(r + c);
            await sheet.AddRowAsync(cells);
        }
        await sheet.FinishAsync();
    }
}

/// <summary>
/// The buffered (whole-workbook DOM) write comparison: 50,000 rows × 10
/// columns of mixed scalars built fully in memory, then saved. 50k keeps
/// ClosedXML's DOM tractable in wall-clock while staying in the same
/// cell-count region as design §4's 30k-in-memory target. SpreadCheetah is
/// excluded — it has no buffered/DOM model, only the streaming path measured
/// in <see cref="ComparativeStreamingWrite"/>.
/// </summary>
[Config(typeof(ComparativeConfig))]
public class ComparativeBufferedWrite
{
    private const int Rows = 50_000;
    private const int Cols = 10;

    private string _tmpPath = null!;

    [GlobalSetup]
    public void Setup() =>
        _tmpPath = Path.Combine(Path.GetTempPath(), $"netxlsx-cmp-buf-{Guid.NewGuid():N}.xlsx");

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_tmpPath)) File.Delete(_tmpPath);
    }

    /// <summary>NetXlsx in-memory write — the baseline.</summary>
    [Benchmark(Baseline = true, Description = "NetXlsx (buffered)")]
    public void NetXlsx_Buffered()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        for (int r = 1; r <= Rows; r++)
        {
            var row = sheet.AppendRow();
            row.Set(1, $"row-{r}").Set(2, r).Set(3, r * 1.5);
            for (int c = 4; c <= Cols; c++) row.Set(c, r + c);
        }
        wb.Save(_tmpPath);
    }

    /// <summary>ClosedXML DOM write of the identical workload.</summary>
    [Benchmark(Description = "ClosedXML")]
    public void ClosedXML_Buffered()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        for (int r = 1; r <= Rows; r++)
        {
            ws.Cell(r, 1).Value = $"row-{r}";
            ws.Cell(r, 2).Value = r;
            ws.Cell(r, 3).Value = r * 1.5;
            for (int c = 4; c <= Cols; c++) ws.Cell(r, c).Value = r + c;
        }
        wb.SaveAs(_tmpPath);
    }
}

/// <summary>
/// The read comparison: open a 50,000-row × 10-column file (authored once by
/// NetXlsx in <see cref="Setup"/> so both readers see identical bytes) and sum
/// one numeric column end to end. SpreadCheetah is excluded — it is write-only.
/// </summary>
[Config(typeof(ComparativeConfig))]
public class ComparativeRead
{
    private const int Rows = 50_000;

    private string _srcPath = null!;

    [GlobalSetup]
    public void Setup()
    {
        _srcPath = Path.Combine(Path.GetTempPath(), $"netxlsx-cmp-read-{Guid.NewGuid():N}.xlsx");
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("Data");
        for (int r = 1; r <= Rows; r++)
        {
            var row = sheet.AppendRow();
            row.Set(1, $"row-{r}").Set(2, r).Set(3, r * 1.5);
        }
        wb.Save(_srcPath);
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_srcPath)) File.Delete(_srcPath);
    }

    /// <summary>NetXlsx open + column sum — the baseline.</summary>
    [Benchmark(Baseline = true, Description = "NetXlsx")]
    public double NetXlsx_Read()
    {
        using var wb = Workbook.Open(_srcPath);
        var sheet = wb["Data"];
        double sum = 0;
        for (int r = 1; r <= Rows; r++) sum += sheet[r, 3].GetNumber() ?? 0;
        return sum;
    }

    /// <summary>ClosedXML open + column sum of the same file.</summary>
    [Benchmark(Description = "ClosedXML")]
    public double ClosedXML_Read()
    {
        using var wb = new XLWorkbook(_srcPath);
        var ws = wb.Worksheet("Data");
        double sum = 0;
        for (int r = 1; r <= Rows; r++) sum += ws.Cell(r, 3).GetDouble();
        return sum;
    }
}
