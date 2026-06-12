// Public entry point. See design §6.1.
// v2.0.0 (I-82): every factory routes to the Open XML SDK engine. The legacy
// NPOI engine is retired; its Xssf*/Sxssf* internals are deleted with the
// NPOI package at the cutover's drop commit.

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

namespace NetXlsx;

/// <summary>
/// Entry point for creating and opening NetXlsx workbooks.
/// </summary>
public static class Workbook
{
    private const int MaxSheetNameLength = 31;
    // Excel's documented forbidden set is / \ ? * : [ ] — the colon was
    // missing here until R-9 (it collides with 3D-reference syntax, so a
    // colon-bearing name corrupts every formula that mentions the sheet).
    // This one array feeds IsValidSheetName, SanitizeSheetName and
    // ValidateSheetName so the three can never disagree.
    private static readonly char[] s_invalidSheetNameChars = { '\\', '/', '?', '*', ':', '[', ']' };
    // Excel also forbids a LEADING or TRAILING apostrophe (interior is
    // fine — "O'Brien" is a legal sheet name) and reserves "History" for
    // shared-workbook change tracking. LibreOffice silently renames an
    // apostrophe-edged sheet to "Sheet1" on resave — accepting one here
    // means data corruption one consumer downstream (R-9).
    private const string ReservedSheetName = "History";

    /// <summary>
    /// Creates a new, empty workbook with no sheets.
    /// </summary>
    /// <param name="options">
    /// Per-workbook configuration (write-side limits, default font,
    /// display culture, date system). When <c>null</c>, uses
    /// <see cref="WorkbookOptions"/> defaults — equivalent to passing
    /// <c>new WorkbookOptions()</c>.
    /// </param>
    public static IWorkbook Create(WorkbookOptions? options = null)
    {
        return OoxmlWorkbook.Create(options ?? new WorkbookOptions());
    }

    /// <summary>
    /// Creates a new, empty macro-enabled (<c>.xlsm</c>) workbook with no
    /// sheets (decision I-69). The resulting workbook is structurally
    /// identical to <see cref="Create"/> but uses the macro-enabled OOXML
    /// content type, so it can carry VBA project parts through open/save
    /// round-trips. NetXlsx does not read, write, or execute VBA — the
    /// macro content is passthrough only.
    /// </summary>
    /// <param name="options">
    /// Per-workbook configuration. When <c>null</c>, uses
    /// <see cref="WorkbookOptions"/> defaults.
    /// </param>
    public static IWorkbook CreateMacroEnabled(WorkbookOptions? options = null)
    {
        return OoxmlWorkbook.Create(
            options ?? new WorkbookOptions(),
            SpreadsheetDocumentType.MacroEnabledWorkbook);
    }

    /// <summary>
    /// Creates a new, empty <b>streaming</b> workbook. Use when writing more
    /// than ~30k rows — past that threshold an in-memory <see cref="Create"/>
    /// workbook exceeds the design's memory budget per spike 2 (see design §5).
    /// Rows stream forward-only through <c>OpenXmlWriter</c> to per-sheet temp
    /// files; memory stays bounded by
    /// <see cref="StreamingOptions.RowAccessWindowSize"/>. <c>Save</c> is
    /// single-shot: the package is assembled once, after which the workbook
    /// rejects further writes (fail-loud).
    /// </summary>
    /// <param name="options">Streaming-specific knobs (row-access
    /// window size, temp-file compression).</param>
    public static IStreamingWorkbook CreateStreaming(StreamingOptions? options = null)
    {
        return new OoxmlStreamingWorkbook(options ?? new StreamingOptions());
    }

    /// <summary>Opens an existing <c>.xlsx</c> or <c>.xlsm</c> workbook from a file path.</summary>
    /// <exception cref="FileNotFoundException">The file does not exist.</exception>
    /// <exception cref="MalformedFileException">The file is not a valid <c>.xlsx</c> / <c>.xlsm</c> workbook.</exception>
    public static IWorkbook Open(string path, WorkbookOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(path);
        var opts = options ?? new WorkbookOptions();

        // FileNotFoundException / DirectoryNotFoundException / UnauthorizedAccessException
        // are all standard I/O exceptions and should propagate verbatim — they are not
        // "the file is malformed" failures, they are "the file isn't accessible" failures.
        // Malformed-input classification (the I-60-equivalent gate, quirk #17)
        // lives in OoxmlWorkbook.OpenCore.
        return OoxmlWorkbook.Open(path, opts);
    }

    /// <summary>Opens an existing <c>.xlsx</c> or <c>.xlsm</c> workbook from a stream.</summary>
    /// <param name="stream">Readable, seekable stream positioned at 0 (decisions #50 / I14).</param>
    /// <param name="leaveOpen">If <c>false</c>, the stream is disposed after the workbook is read. Default <c>true</c> per BCL convention.</param>
    /// <exception cref="ArgumentException">The stream is not readable or not seekable.</exception>
    /// <exception cref="MalformedFileException">The stream content is not a valid <c>.xlsx</c> / <c>.xlsm</c> workbook.</exception>
    public static IWorkbook Open(Stream stream, bool leaveOpen = true, WorkbookOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        // The CanSeek + Position guards predate the engine swap and are KEPT at
        // the cutover (decided 2026-06-04): change only what must change.
        // Relaxing to readable-only is additive and tracked as a deliberate
        // post-v2.0.0 enhancement, not a rider on the flip.
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        if (!stream.CanSeek) throw new ArgumentException("Stream must be seekable (decisions #50 / I14).", nameof(stream));
        if (stream.Position != 0) throw new ArgumentException("Stream must be positioned at 0.", nameof(stream));
        var opts = options ?? new WorkbookOptions();
        return OoxmlWorkbook.Open(stream, leaveOpen, opts);
    }

    /// <summary>Asynchronously opens an existing <c>.xlsx</c> workbook from a path.</summary>
    public static Task<IWorkbook> OpenAsync(string path, WorkbookOptions? options = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(path);
        // The open is synchronous; we offload to the thread pool per decision
        // #30 / §7.1. CancellationToken is honored only before the offload
        // begins; mid-open cancellation is not supported.
        ct.ThrowIfCancellationRequested();
        return Task.Run(() => Open(path, options), ct);
    }

    /// <summary>Asynchronously opens an existing <c>.xlsx</c> workbook from a stream.</summary>
    public static Task<IWorkbook> OpenAsync(Stream stream, bool leaveOpen = true, WorkbookOptions? options = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ct.ThrowIfCancellationRequested();
        return Task.Run(() => Open(stream, leaveOpen, options), ct);
    }

    /// <summary>
    /// Returns <c>true</c> if <paramref name="proposed"/> meets Excel's
    /// sheet-name rules: length 1..31, no <c>\ / ? * : [ ]</c> characters,
    /// no leading or trailing apostrophe, and not the reserved name
    /// <c>History</c> (case-insensitive). Does not check workbook-level
    /// uniqueness.
    /// </summary>
    public static bool IsValidSheetName(string proposed)
    {
        if (string.IsNullOrEmpty(proposed)) return false;
        if (proposed.Length > MaxSheetNameLength) return false;
        if (proposed.IndexOfAny(s_invalidSheetNameChars) >= 0) return false;
        if (proposed[0] == '\'' || proposed[^1] == '\'') return false;
        if (string.Equals(proposed, ReservedSheetName, StringComparison.OrdinalIgnoreCase)) return false;
        return true;
    }

    /// <summary>
    /// Returns a sanitized version of <paramref name="proposed"/> that
    /// satisfies <see cref="IsValidSheetName"/>: truncated to 31 chars,
    /// invalid characters replaced with underscore. Empty input becomes
    /// <c>"Sheet"</c>.
    /// <para>
    /// <b>Important:</b> this method does <b>not</b> guarantee uniqueness
    /// against an existing workbook. Sanitization can produce collisions
    /// (e.g. <c>"Q1/2026"</c> and <c>"Q1?2026"</c> both sanitize to
    /// <c>"Q1_2026"</c>). When you need both sanitization <i>and</i>
    /// uniqueness against an existing workbook, call
    /// <see cref="SuggestSheetName"/> instead — it sanitizes internally
    /// and then resolves collisions with a numeric suffix.
    /// </para>
    /// </summary>
    public static string SanitizeSheetName(string proposed)
    {
        if (string.IsNullOrEmpty(proposed)) return "Sheet";
        var chars = proposed.ToCharArray();
        for (int i = 0; i < chars.Length; i++)
        {
            if (Array.IndexOf(s_invalidSheetNameChars, chars[i]) >= 0)
                chars[i] = '_';
        }
        if (chars[0] == '\'') chars[0] = '_';
        var sanitized = new string(chars);
        if (sanitized.Length > MaxSheetNameLength)
            sanitized = sanitized.Substring(0, MaxSheetNameLength);
        // The trailing-apostrophe fix runs AFTER truncation — truncating a
        // longer name can expose a new last character.
        if (sanitized[^1] == '\'')
            sanitized = string.Concat(sanitized.AsSpan(0, sanitized.Length - 1), "_");
        // Reserved name: exact (case-insensitive) match only — appending an
        // underscore stays within the length limit ("History" is 7 chars).
        if (string.Equals(sanitized, ReservedSheetName, StringComparison.OrdinalIgnoreCase))
            sanitized += "_";
        return sanitized;
    }

    /// <summary>
    /// Returns <paramref name="proposed"/> if no sheet with that name
    /// exists in <paramref name="workbook"/> (case-insensitive); otherwise
    /// returns <c>"<paramref name="proposed"/> (2)"</c>, <c>"(3)"</c>, etc.
    /// until an unused name is found. Truncates to the 31-character limit
    /// while preserving the disambiguating suffix.
    /// <para>
    /// The numeric-suffix search caps at 10,000 attempts (numbers
    /// <c>(2)</c>..<c>(9999)</c>). At that point — essentially unreachable
    /// under any realistic workbook — the method falls through to a
    /// GUID-tagged name (8-hex suffix appended to a truncated base)
    /// rather than throwing. This abandons the <c>(N)</c> pattern at the
    /// ceiling but guarantees the method has a defined return for any
    /// input. Excel's own 31-char limit makes the ceiling far easier to
    /// reach with deliberately pathological input than with real data.
    /// </para>
    /// </summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    public static string SuggestSheetName(IWorkbook workbook, string proposed)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(proposed);

        var sanitized = SanitizeSheetName(proposed);
        if (!workbook.TryGetSheet(sanitized, out _))
            return sanitized;

        for (int suffix = 2; suffix < 10_000; suffix++)
        {
            var suffixStr = $" ({suffix.ToString(System.Globalization.CultureInfo.InvariantCulture)})";
            var maxBaseLen = MaxSheetNameLength - suffixStr.Length;
            var baseName = sanitized.Length > maxBaseLen
                ? sanitized.Substring(0, maxBaseLen)
                : sanitized;
            var candidate = baseName + suffixStr;
            if (!workbook.TryGetSheet(candidate, out _))
                return candidate;
        }

        // Effectively unreachable under any realistic workbook (10k+
        // identical sheet-name attempts). Fall through to a GUID-tagged
        // candidate so the method has a defined exit even at the limit.
        var fallbackTag = Guid.NewGuid().ToString("N").Substring(0, 8);
        var fallbackBase = sanitized.Length > MaxSheetNameLength - 9
            ? sanitized.Substring(0, MaxSheetNameLength - 9)
            : sanitized;
        return fallbackBase + "_" + fallbackTag;
    }

    internal static void ValidateSheetName(string name)
    {
        if (name is null) throw new SheetNameException("<null>", "sheet name is null");
        if (name.Length == 0) throw new SheetNameException(name, "sheet name is empty");
        if (name.Length > MaxSheetNameLength)
            throw new SheetNameException(name, $"sheet name exceeds {MaxSheetNameLength} characters");
        int invalidIdx = name.IndexOfAny(s_invalidSheetNameChars);
        if (invalidIdx >= 0)
            throw new SheetNameException(name, $"sheet name contains invalid character '{name[invalidIdx]}' at position {invalidIdx}");
        if (name[0] == '\'' || name[^1] == '\'')
            throw new SheetNameException(name, "sheet name cannot begin or end with an apostrophe (Excel rule; LibreOffice silently renames such sheets)");
        if (string.Equals(name, ReservedSheetName, StringComparison.OrdinalIgnoreCase))
            throw new SheetNameException(name, "'History' is reserved by Excel for shared-workbook change tracking");
    }
}
