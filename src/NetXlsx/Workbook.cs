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
        try
        {
            // NPOI's XSSFWorkbook(string) opens read-only; for read-write we
            // pass a FileStream we own to ensure write-back works at Save().
            var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            try
            {
                var underlying = new XSSFWorkbook(fs);
                return new XssfWorkbook(underlying);
            }
            catch (Exception ex)
            {
                fs.Dispose();
                throw new MalformedFileException($"Failed to open '{path}' as .xlsx", ex);
            }
            finally
            {
                // NPOI copies the stream content into memory; we can release ours.
                fs.Dispose();
            }
        }
        catch (FileNotFoundException)
        {
            throw;
        }
        catch (MalformedFileException)
        {
            throw;
        }
        catch (Exception ex) when (ex is not WorkbookException)
        {
            throw new MalformedFileException($"Failed to open '{path}' as .xlsx", ex);
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
        catch (Exception ex) when (ex is not WorkbookException)
        {
            throw new MalformedFileException("Stream content is not a valid .xlsx workbook.", ex);
        }
        finally
        {
            if (!leaveOpen) stream.Dispose();
        }
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
