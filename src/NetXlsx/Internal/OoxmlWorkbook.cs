// I-82 engine swap — Open XML SDK-backed IWorkbook.
//
// This is the v2.0.0 engine, grown additively alongside the NPOI-backed
// XssfWorkbook under the parallel-engine / late-cutover strategy (see design
// I-82 and the continuation plan's "Engine swap strategy"). It is reached only
// through Workbook.CreateOoxml() / Workbook.OpenOoxml(...); the default
// Workbook.Create() / Open() still return the NPOI engine until the cutover.
//
// Foundation slice scope (this commit): Create / Open / Save / Dispose,
// AddSheet, sheet enumeration + indexers, and the OpenXmlDocument escape hatch.
// Everything else throws NotYet(...) and lands slice by slice (cells & rows ->
// styles -> rich text -> merges/panes/grouping -> drawings -> CF/validation/
// tables/autofilter/sort -> charts -> streaming).
//
// Engine model: the workbook owns an in-memory MemoryStream and a
// SpreadsheetDocument opened read/write over it. This mirrors the NPOI wrapper's
// "load into memory, save anywhere, repeatedly" semantics — Open copies the
// source into the owned buffer (the caller's file/stream is never mutated), and
// Save clones the live package into a throwaway buffer (finalizing the zip
// central directory) before copying the finished bytes to the target. Cloning
// to a private buffer sidesteps any ambiguity about whether the SDK's
// Clone(Stream) takes ownership of the destination stream.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlWorkbook : IWorkbook
{
    private readonly SpreadsheetDocument _document;
    private readonly MemoryStream _backing;
    private readonly WorkbookOptions _options;
    private readonly Dictionary<string, OoxmlSheet> _sheetsByName = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<OoxmlSheet> _sheetsByIndex = new();
    private bool _disposed;

    internal WorkbookOptions Options => _options;

    private OoxmlWorkbook(SpreadsheetDocument document, MemoryStream backing, WorkbookOptions options)
    {
        _document = document;
        _backing = backing;
        _options = options;
    }

    // ---- Factories (called by Workbook.CreateOoxml / OpenOoxml) -------------

    internal static OoxmlWorkbook Create(WorkbookOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        var backing = new MemoryStream();
        try
        {
            var document = SpreadsheetDocument.Create(backing, SpreadsheetDocumentType.Workbook);
            var wbPart = document.AddWorkbookPart();
            wbPart.Workbook = new S.Workbook();
            wbPart.Workbook.AppendChild(new S.Sheets());
            // WorkbookOptions (default font, 1904 date system) are applied in the
            // styles / cells slices, where the stylesheet and workbookPr parts are
            // modeled. Stored now so those slices have them; not applied yet.
            return new OoxmlWorkbook(document, backing, options);
        }
        catch
        {
            backing.Dispose();
            throw;
        }
    }

    internal static OoxmlWorkbook Open(string path, WorkbookOptions options)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentNullException.ThrowIfNull(options);
        // FileNotFound / DirectoryNotFound / UnauthorizedAccess propagate verbatim:
        // they are "not accessible" failures, not "malformed" failures.
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return OpenCore(fs, $"Failed to open '{path}' as .xlsx", options);
    }

    internal static OoxmlWorkbook Open(Stream stream, bool leaveOpen, WorkbookOptions options)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(options);
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        try
        {
            return OpenCore(stream, "Stream content is not a valid .xlsx workbook.", options);
        }
        finally
        {
            if (!leaveOpen) stream.Dispose();
        }
    }

    private static OoxmlWorkbook OpenCore(Stream source, string malformedMessage, WorkbookOptions options)
    {
        // Copy the source into an owned, editable buffer so Save can target any
        // destination and the caller's stream is never mutated.
        var backing = new MemoryStream();
        try
        {
            source.CopyTo(backing);
            backing.Position = 0;
        }
        catch
        {
            backing.Dispose();
            throw;
        }

        SpreadsheetDocument document;
        try
        {
            document = SpreadsheetDocument.Open(backing, isEditable: true);
        }
        catch (Exception ex) when (IsKnownMalformedOpenException(ex))
        {
            backing.Dispose();
            throw new MalformedFileException(malformedMessage, ex);
        }
        catch
        {
            backing.Dispose();
            throw;
        }

        var wb = new OoxmlWorkbook(document, backing, options);
        wb.IndexExistingSheets();
        return wb;
    }

    // Open XML SDK surfaces malformed packages as OpenXmlPackageException, the
    // System.IO.Packaging FileFormatException, or low-level IO/XML/format
    // failures. Critical runtime faults and our own exceptions propagate verbatim.
    private static bool IsKnownMalformedOpenException(Exception ex)
    {
        if (ex is WorkbookException) return false;
        if (ex is OutOfMemoryException or StackOverflowException or OperationCanceledException) return false;
        if (ex is ArgumentNullException) return false;

        var typeName = ex.GetType().FullName ?? string.Empty;
        if (typeName.StartsWith("DocumentFormat.OpenXml.", StringComparison.Ordinal)) return true;

        return ex is System.IO.InvalidDataException
            or System.IO.FileFormatException
            or System.IO.IOException
            or System.Xml.XmlException
            or System.ArgumentException
            or System.FormatException;
    }

    private void IndexExistingSheets()
    {
        var sheets = _document.WorkbookPart?.Workbook?.GetFirstChild<S.Sheets>();
        if (sheets is null) return;
        var wbPart = _document.WorkbookPart!;
        foreach (var sheet in sheets.Elements<S.Sheet>())
        {
            var name = sheet.Name?.Value ?? string.Empty;
            // Resolve the worksheet part backing this sheet via its r:id.
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!);
            var wrapper = new OoxmlSheet(this, name, wsPart);
            _sheetsByIndex.Add(wrapper);
            _sheetsByName[name] = wrapper;
        }
    }

    // ---- Bones --------------------------------------------------------------

    public int SheetCount
    {
        get { ThrowIfDisposed(); return _sheetsByIndex.Count; }
    }

    public ISheet this[string name]
    {
        get
        {
            ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(name);
            if (_sheetsByName.TryGetValue(name, out var sheet)) return sheet;
            throw new KeyNotFoundException($"Sheet '{name}' not found.");
        }
    }

    public ISheet this[int index]
    {
        get
        {
            ThrowIfDisposed();
            if (index < 0 || index >= _sheetsByIndex.Count)
                throw new ArgumentOutOfRangeException(nameof(index), index, $"Sheet index must be in [0, {_sheetsByIndex.Count - 1}].");
            return _sheetsByIndex[index];
        }
    }

    public ISheet AddSheet(string name)
    {
        ThrowIfDisposed();
        Workbook.ValidateSheetName(name);
        if (_sheetsByName.ContainsKey(name))
            throw new SheetNameException(name, "a sheet with this name already exists (case-insensitive)");

        var wbPart = _document.WorkbookPart
            ?? throw new InvalidOperationException("Workbook has no workbook part.");
        var workbook = wbPart.Workbook
            ?? throw new InvalidOperationException("Workbook part has no workbook element.");
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        wsPart.Worksheet = new S.Worksheet(new S.SheetData());

        var sheets = workbook.GetFirstChild<S.Sheets>()
            ?? workbook.AppendChild(new S.Sheets());
        var sheetElement = new S.Sheet
        {
            Id = wbPart.GetIdOfPart(wsPart),
            SheetId = NextSheetId(sheets),
            Name = name,
        };
        sheets.AppendChild(sheetElement);

        var wrapper = new OoxmlSheet(this, name, wsPart);
        _sheetsByIndex.Add(wrapper);
        _sheetsByName[name] = wrapper;
        return wrapper;
    }

    private static uint NextSheetId(S.Sheets sheets)
    {
        uint max = 0;
        foreach (var sheet in sheets.Elements<S.Sheet>())
        {
            if (sheet.SheetId?.Value is uint id && id > max) max = id;
        }
        return max + 1;
    }

    public bool TryGetSheet(string name, [MaybeNullWhen(false)] out ISheet sheet)
    {
        ThrowIfDisposed();
        if (name is null) { sheet = null; return false; }
        if (_sheetsByName.TryGetValue(name, out var x)) { sheet = x; return true; }
        sheet = null;
        return false;
    }

    public void Save(Stream stream, bool leaveOpen = true)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));

        // Flush the strongly-typed DOM into the in-memory package parts, then
        // clone into a throwaway buffer whose disposal finalizes the zip central
        // directory. The live document stays open and re-saveable.
        if (_document.CanSave) _document.Save();
        using (var tmp = new MemoryStream())
        {
            using (_document.Clone(tmp)) { }
            tmp.Position = 0;
            tmp.CopyTo(stream);
        }

        if (!leaveOpen) stream.Dispose();
    }

    public void Save(string path)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(path);
        using var fs = File.Create(path);
        Save(fs, leaveOpen: false);
    }

    public Task SaveAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default)
    {
        ThrowIfDisposed();
        ct.ThrowIfCancellationRequested();
        return Task.Run(() => Save(stream, leaveOpen), ct);
    }

    public Task SaveAsync(string path, CancellationToken ct = default)
    {
        ThrowIfDisposed();
        ct.ThrowIfCancellationRequested();
        return Task.Run(() => Save(path), ct);
    }

    // Escape hatch (I-82): the SDK document is the OOXML engine's escape hatch.
    public SpreadsheetDocument? OpenXmlDocument
    {
        get { ThrowIfDisposed(); return _document; }
    }

    // No NPOI workbook exists on the SDK engine; the NPOI escape hatch diverges.
    public NPOI.XSSF.UserModel.XSSFWorkbook Underlying => throw new NotSupportedException(
        "IWorkbook.Underlying (NPOI XSSFWorkbook) is not available on the Open XML " +
        "SDK engine (I-82). Use IWorkbook.OpenXmlDocument for the SDK escape hatch.");

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _document.Dispose();
        _backing.Dispose();
    }

    internal void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(IWorkbook));
    }

    // ---- Not-yet-implemented surface (lands slice by slice; see I-82) -------

    private static NotImplementedException NotYet([CallerMemberName] string? member = null)
        => new(
            $"IWorkbook.{member} is not yet implemented on the Open XML SDK engine " +
            "(I-82 engine swap). It lands in a later slice; until then use the " +
            "legacy engine (Workbook.Create/Open) for this operation, or track the " +
            "swap in docs/design.md (I-82).");

    public INamedRange AddNamedRange(string name, string formula, string? sheetScope = null) => throw NotYet();
    public IReadOnlyList<INamedRange> NamedRanges => throw NotYet();

    public StylePoolDiagnostics GetStylePoolDiagnostics() => throw NotYet();
    public void RegisterStyle(string name, CellStyle style) => throw NotYet();
    public CellStyle? GetRegisteredStyle(string name) => throw NotYet();
    public IReadOnlyCollection<string> RegisteredStyleNames => throw NotYet();

    public void Protect(WorkbookProtection? options = null) => throw NotYet();
    public void ProtectWithPassword(string password, WorkbookProtection? options = null) => throw NotYet();
    public void Unprotect() => throw NotYet();
    public bool IsProtected => throw NotYet();

    public void SetThemeXml(byte[] themeXml) => throw NotYet();
    public byte[]? GetThemeXml() => throw NotYet();
    public Color? ResolveThemeColor(int index, double tint = 0) => throw NotYet();
    public Color? ResolveThemeColor(ThemeColor color) => throw NotYet();
    public Color? ResolveThemeColor(string schemeName, double tint = 0) => throw NotYet();
    public int? GetThemeLineWidthEmu(int oneBasedIdx) => throw NotYet();

    public bool IsMacroEnabled => throw NotYet();
}
