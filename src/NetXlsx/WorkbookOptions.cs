using System.Globalization;

namespace NetXlsx;

/// <summary>
/// Per-workbook configuration. Passed to <see cref="Workbook.Create"/>,
/// <see cref="Workbook.Open(string, WorkbookOptions?)"/>, and (via the
/// derived <see cref="StreamingOptions"/>) to
/// <see cref="Workbook.CreateStreaming"/>.
/// <para>
/// All fields are init-only — construct once at workbook creation and
/// pass through. Defaults match Excel and design §6.1.
/// </para>
/// </summary>
public class WorkbookOptions
{
    /// <summary>
    /// Culture used by <see cref="ICell.GetString"/> when rendering
    /// number/date cells to their displayed form. Reading raw typed
    /// values is always invariant per design §7.2. Default:
    /// <see cref="CultureInfo.InvariantCulture"/>.
    /// </summary>
    public CultureInfo DisplayCulture { get; init; } = CultureInfo.InvariantCulture;

    /// <summary>
    /// Whether the workbook uses the 1900 or 1904 date system.
    /// Workbooks authored on Mac Excel pre-2016 default to 1904; all
    /// other Excel installations default to 1900. Default:
    /// <see cref="DateSystem.Excel1900"/>.
    /// </summary>
    public DateSystem DateSystem { get; init; } = DateSystem.Excel1900;

    // ---- Read-side safety (zip-bomb defense; not applied on write) ----

    /// <summary>
    /// Maximum decompressed payload size accepted by
    /// <see cref="Workbook.Open(string, WorkbookOptions?)"/>. Mitigates
    /// zip-bomb attacks. Default: 256 MiB.
    /// </summary>
    public long ReadMaxUncompressedBytes { get; init; } = 256L * 1024 * 1024;

    /// <summary>
    /// Maximum number of sheets accepted when opening a workbook.
    /// Default: 1000.
    /// </summary>
    public int ReadMaxSheets { get; init; } = 1000;

    // ---- Excel hard limits (write-side enforcement) -------------------

    /// <summary>
    /// Cap on rows per sheet enforced on write. Default is Excel's
    /// hard cap (1,048,576).
    /// </summary>
    public int MaxRowsPerSheet { get; init; } = 1_048_576;

    /// <summary>
    /// Cap on columns per sheet enforced on write. Default is Excel's
    /// hard cap (16,384, the "XFD" column).
    /// </summary>
    public int MaxColsPerSheet { get; init; } = 16_384;

    /// <summary>
    /// Cap on per-cell text length enforced on write. Default is
    /// Excel's hard cap (32,767 characters).
    /// </summary>
    public int MaxCellTextLength { get; init; } = 32_767;

    /// <summary>
    /// Default font family applied to new cells when no explicit
    /// font is set. Default: <c>"Calibri"</c>.
    /// </summary>
    public string DefaultFontName { get; init; } = "Calibri";

    /// <summary>
    /// Default font size in points. Default: <c>11</c>.
    /// </summary>
    public double DefaultFontSize { get; init; } = 11;

    /// <summary>
    /// Opt-in strict concurrency detection (decision I-59). When
    /// <c>true</c>, every structural mutating path on <see cref="IWorkbook"/>
    /// takes a real per-workbook lock, eliminating the gap between the
    /// default opportunistic reentry counter (decision #43) and silent
    /// corruption from concurrent threads.
    /// <para>
    /// Default <c>false</c> — single-threaded callers don't pay the
    /// lock cost. Set this when the workbook may be touched from
    /// multiple threads, even with external serialization, and you
    /// want a hard guarantee that concurrent mutations surface as
    /// <see cref="System.InvalidOperationException"/> rather than
    /// silently corrupting state.
    /// </para>
    /// <para>
    /// Strict mode does not make the workbook thread-safe for
    /// reads — concurrent reads of any kind remain undefined.
    /// </para>
    /// </summary>
    public bool StrictConcurrencyDetection { get; init; }
}

/// <summary>
/// Options for <see cref="Workbook.CreateStreaming"/>. Extends
/// <see cref="WorkbookOptions"/> with streaming-specific knobs.
/// </summary>
public sealed class StreamingOptions : WorkbookOptions
{
    /// <summary>
    /// Number of rows held in memory before older rows are flushed to
    /// disk. Passed through to NPOI's
    /// <c>SXSSFWorkbook(XSSFWorkbook, int)</c> constructor. Larger
    /// values reduce flush frequency but raise peak memory. Default:
    /// <c>100</c> (NPOI default).
    /// </summary>
    public int RowAccessWindowSize { get; init; } = 100;

    /// <summary>
    /// Whether SXSSF compresses on-disk temp files. Trades CPU for
    /// disk I/O — usually a wash for typical row sizes. Default:
    /// <c>false</c>.
    /// </summary>
    public bool CompressTempFiles { get; init; }
}

/// <summary>The Excel date epoch.</summary>
public enum DateSystem
{
    /// <summary>The 1900-based epoch — Excel's standard default.</summary>
    Excel1900,
    /// <summary>The 1904-based epoch — Mac Excel pre-2016 default.</summary>
    Excel1904,
}
