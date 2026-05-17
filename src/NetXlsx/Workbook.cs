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

    /// <summary>Creates a new, empty workbook with no sheets.</summary>
    public static IWorkbook Create()
    {
        var underlying = new XSSFWorkbook();
        return new XssfWorkbook(underlying);
    }

    /// <summary>Opens an existing <c>.xlsx</c> workbook from a file path.</summary>
    /// <exception cref="FileNotFoundException">The file does not exist.</exception>
    /// <exception cref="MalformedFileException">The file is not a valid <c>.xlsx</c> workbook.</exception>
    public static IWorkbook Open(string path)
    {
        ArgumentNullException.ThrowIfNull(path);

        // FileNotFoundException / DirectoryNotFoundException / UnauthorizedAccessException
        // are all standard I/O exceptions and should propagate verbatim — they are not
        // "the file is malformed" failures, they are "the file isn't accessible" failures.
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        try
        {
            var underlying = new XSSFWorkbook(fs);
            return new XssfWorkbook(underlying);
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

    /// <summary>Opens an existing <c>.xlsx</c> workbook from a stream.</summary>
    /// <param name="stream">Readable, seekable stream positioned at 0 (decisions #50 / I14).</param>
    /// <param name="leaveOpen">If <c>false</c>, the stream is disposed after the workbook is read. Default <c>true</c> per BCL convention.</param>
    /// <exception cref="ArgumentException">The stream is not readable or not seekable.</exception>
    /// <exception cref="MalformedFileException">The stream content is not a valid <c>.xlsx</c> workbook.</exception>
    public static IWorkbook Open(Stream stream, bool leaveOpen = true)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        if (!stream.CanSeek) throw new ArgumentException("Stream must be seekable (NPOI requires seek).", nameof(stream));
        if (stream.Position != 0) throw new ArgumentException("Stream must be positioned at 0.", nameof(stream));

        try
        {
            var underlying = new XSSFWorkbook(stream);
            return new XssfWorkbook(underlying);
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
        return ex is System.IO.InvalidDataException
            or System.IO.IOException
            or System.IO.EndOfStreamException
            or System.Xml.XmlException
            or FormatException;
    }

    /// <summary>Asynchronously opens an existing <c>.xlsx</c> workbook from a path.</summary>
    public static Task<IWorkbook> OpenAsync(string path, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(path);
        // NPOI is synchronous; we offload to the thread pool per decision #30 / §7.1.
        // CancellationToken is honored only before the offload begins; mid-NPOI
        // cancellation is not supported.
        ct.ThrowIfCancellationRequested();
        return Task.Run(() => Open(path), ct);
    }

    /// <summary>Asynchronously opens an existing <c>.xlsx</c> workbook from a stream.</summary>
    public static Task<IWorkbook> OpenAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ct.ThrowIfCancellationRequested();
        return Task.Run(() => Open(stream, leaveOpen), ct);
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
