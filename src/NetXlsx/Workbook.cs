// Public entry point. See design §6.1.
// v0.2.0 vertical-slice subset: Create / Open / OpenAsync + sheet-name validation.

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

/// <summary>
/// Entry point for creating and opening NetXlsx workbooks.
/// </summary>
public static class Workbook
{
    private const int MaxSheetNameLength = 31;
    private static readonly char[] s_invalidSheetNameChars = { '\\', '/', '?', '*', '[', ']' };

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
        var underlying = new XSSFWorkbook();
        return new XssfWorkbook(underlying, options ?? new WorkbookOptions());
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
        var underlying = new XSSFWorkbook(XSSFWorkbookType.XLSM);
        return new XssfWorkbook(underlying, options ?? new WorkbookOptions());
    }

    /// <summary>
    /// Creates a new, empty <b>streaming</b> workbook backed by NPOI's
    /// SXSSF writer. Use when writing more than ~30k rows — past that
    /// threshold an in-memory <see cref="Create"/> workbook exceeds the
    /// design's memory budget per spike 2 (see design §5).
    /// </summary>
    /// <param name="options">Streaming-specific knobs (row-access
    /// window size, temp-file compression). Defaults match NPOI.</param>
    public static IStreamingWorkbook CreateStreaming(StreamingOptions? options = null)
    {
        return new SxssfWorkbook(options ?? new StreamingOptions());
    }

    // ---- I-82 engine swap: Open XML SDK factories --------------------------
    // These return the new Open XML SDK-backed engine, grown additively
    // alongside the default NPOI engine under the parallel-engine / late-cutover
    // strategy (design I-82). The default Create()/Open() keep returning the NPOI
    // engine until the v2.0.0 cutover slice, which retires NPOI and folds the SDK
    // engine into Create()/Open() directly. A workbook from these factories
    // reports a non-null IWorkbook.OpenXmlDocument and throws on the NPOI
    // IWorkbook.Underlying escape hatch.

    /// <summary>
    /// Creates a new, empty workbook on the Open XML SDK engine (decision
    /// I-82). Counterpart to <see cref="Create"/>, which uses the legacy NPOI
    /// engine. The SDK engine is the v2.0.0 direction; during the swap it is
    /// reached only through this factory.
    /// </summary>
    /// <param name="options">
    /// Per-workbook configuration. When <c>null</c>, uses
    /// <see cref="WorkbookOptions"/> defaults.
    /// </param>
    public static IWorkbook CreateOoxml(WorkbookOptions? options = null)
    {
        return OoxmlWorkbook.Create(options ?? new WorkbookOptions());
    }

    /// <summary>
    /// Opens an existing <c>.xlsx</c> / <c>.xlsm</c> workbook on the Open XML
    /// SDK engine from a file path (decision I-82). Counterpart to
    /// <see cref="Open(string, WorkbookOptions?)"/>, which uses the legacy NPOI
    /// engine.
    /// </summary>
    /// <exception cref="FileNotFoundException">The file does not exist.</exception>
    /// <exception cref="MalformedFileException">The file is not a valid <c>.xlsx</c> / <c>.xlsm</c> workbook.</exception>
    public static IWorkbook OpenOoxml(string path, WorkbookOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(path);
        return OoxmlWorkbook.Open(path, options ?? new WorkbookOptions());
    }

    /// <summary>
    /// Opens an existing <c>.xlsx</c> / <c>.xlsm</c> workbook on the Open XML
    /// SDK engine from a stream (decision I-82). The stream's content is copied
    /// into an owned in-memory buffer, so the stream need only be readable; it
    /// is not mutated.
    /// </summary>
    /// <param name="stream">A readable stream positioned at the workbook's start.</param>
    /// <param name="leaveOpen">If <c>false</c>, the stream is disposed after the workbook is read. Default <c>true</c>.</param>
    /// <exception cref="ArgumentException">The stream is not readable.</exception>
    /// <exception cref="MalformedFileException">The stream content is not a valid workbook.</exception>
    public static IWorkbook OpenOoxml(Stream stream, bool leaveOpen = true, WorkbookOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        return OoxmlWorkbook.Open(stream, leaveOpen, options ?? new WorkbookOptions());
    }

    /// <summary>
    /// Creates a new, empty <b>streaming</b> workbook on the Open XML SDK engine
    /// (decision I-82) — the counterpart to <see cref="CreateStreaming"/>, which
    /// uses the legacy NPOI SXSSF writer. Rows stream forward-only to per-sheet
    /// temp files through <c>OpenXmlWriter</c>; memory stays bounded by
    /// <see cref="StreamingOptions.RowAccessWindowSize"/>. <c>Save</c> is
    /// single-shot: the worksheet XML is finalized and the package assembled
    /// once, after which the workbook rejects further writes (fail-loud, where
    /// NPOI silently discards them). The NPOI escape hatches
    /// (<see cref="IStreamingWorkbook.Underlying"/> /
    /// <see cref="IStreamingSheet.Underlying"/>) throw
    /// <see cref="NotSupportedException"/> on this engine.
    /// </summary>
    /// <param name="options">Streaming-specific knobs (row-access window size,
    /// temp-file compression). Defaults match <see cref="CreateStreaming"/>.</param>
    public static IStreamingWorkbook CreateStreamingOoxml(StreamingOptions? options = null)
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
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        try
        {
            var underlying = new XSSFWorkbook(fs);
            return new XssfWorkbook(underlying, opts);
        }
        catch (Exception ex) when (IsKnownMalformedOpenException(ex))
        {
            throw new MalformedFileException($"Failed to open '{path}' as .xlsx", ex);
        }
        // Other exceptions (OOM, StackOverflow, ArgumentException from BCL,
        // OperationCanceledException, etc.) propagate verbatim. They indicate
        // bugs, resource exhaustion, or programmer error — not malformed input.
        finally
        {
            // NPOI copies stream content into memory; we can release ours.
            fs.Dispose();
        }
    }

    /// <summary>Opens an existing <c>.xlsx</c> or <c>.xlsm</c> workbook from a stream.</summary>
    /// <param name="stream">Readable, seekable stream positioned at 0 (decisions #50 / I14).</param>
    /// <param name="leaveOpen">If <c>false</c>, the stream is disposed after the workbook is read. Default <c>true</c> per BCL convention.</param>
    /// <exception cref="ArgumentException">The stream is not readable or not seekable.</exception>
    /// <exception cref="MalformedFileException">The stream content is not a valid <c>.xlsx</c> / <c>.xlsm</c> workbook.</exception>
    public static IWorkbook Open(Stream stream, bool leaveOpen = true, WorkbookOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        if (!stream.CanSeek) throw new ArgumentException("Stream must be seekable (NPOI requires seek).", nameof(stream));
        if (stream.Position != 0) throw new ArgumentException("Stream must be positioned at 0.", nameof(stream));
        var opts = options ?? new WorkbookOptions();

        try
        {
            var underlying = new XSSFWorkbook(stream);
            return new XssfWorkbook(underlying, opts);
        }
        catch (Exception ex) when (IsKnownMalformedOpenException(ex))
        {
            throw new MalformedFileException("Stream content is not a valid .xlsx workbook.", ex);
        }
        // Other exceptions (OOM, StackOverflow, BCL ArgumentException,
        // OperationCanceledException, etc.) propagate verbatim.
        finally
        {
            if (!leaveOpen) stream.Dispose();
        }
    }

    /// <summary>
    /// Filter used to decide whether an exception thrown by NPOI's
    /// <c>XSSFWorkbook(Stream)</c> constructor represents a malformed
    /// input. Whitelists the NPOI / OPC / IO exception types that can
    /// actually mean "this is not a valid .xlsx." Excludes
    /// <see cref="OutOfMemoryException"/>, <see cref="StackOverflowException"/>,
    /// <see cref="OperationCanceledException"/>, and everything else that
    /// indicates a runtime / programmer fault rather than bad data.
    /// </summary>
    private static bool IsKnownMalformedOpenException(Exception ex)
    {
        // Don't classify our own exceptions as malformed-input.
        if (ex is WorkbookException) return false;
        // Critical runtime exceptions propagate verbatim.
        if (ex is OutOfMemoryException or StackOverflowException or OperationCanceledException) return false;

        // NPOI throws these for OOXML / OPC / underlying-zip failures.
        // Identified by namespace string to avoid taking a hard reference
        // on every internal NPOI exception type (some are nested).
        var typeName = ex.GetType().FullName ?? string.Empty;
        if (typeName.StartsWith("NPOI.", StringComparison.Ordinal)) return true;
        if (typeName.StartsWith("ICSharpCode.SharpZipLib.", StringComparison.Ordinal)) return true;

        // BCL exceptions consistent with bad-input on this code path.
        // The IndexOutOfRange / NullReference / Overflow / Argument*
        // additions are post-v1.0 hardening (decision I-60) — surfaced
        // by the fuzz harness when NPOI's parsers index into truncated
        // arrays / dereference uninitialized parts / overflow length
        // computations on adversarial input. These are still bugs in
        // NPOI ideally, but on the open path the right user-visible
        // contract is "this file is malformed", not a leaked runtime
        // exception. Captured in implementation-notes.md.
        return ex is System.IO.InvalidDataException
            or System.IO.IOException
            or System.IO.EndOfStreamException
            or System.Xml.XmlException
            or System.IndexOutOfRangeException
            or System.NullReferenceException
            or System.OverflowException
            or System.ArgumentOutOfRangeException
            or FormatException;
    }

    /// <summary>Asynchronously opens an existing <c>.xlsx</c> workbook from a path.</summary>
    public static Task<IWorkbook> OpenAsync(string path, WorkbookOptions? options = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(path);
        // NPOI is synchronous; we offload to the thread pool per decision #30 / §7.1.
        // CancellationToken is honored only before the offload begins; mid-NPOI
        // cancellation is not supported.
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
    /// sheet-name rules: length 1..31, no <c>\ / ? * [ ]</c> characters.
    /// Does not check workbook-level uniqueness.
    /// </summary>
    public static bool IsValidSheetName(string proposed)
    {
        if (string.IsNullOrEmpty(proposed)) return false;
        if (proposed.Length > MaxSheetNameLength) return false;
        if (proposed.IndexOfAny(s_invalidSheetNameChars) >= 0) return false;
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
        var sanitized = new string(chars);
        if (sanitized.Length > MaxSheetNameLength)
            sanitized = sanitized.Substring(0, MaxSheetNameLength);
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
    }
}
