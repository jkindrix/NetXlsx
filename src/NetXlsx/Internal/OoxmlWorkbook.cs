// I-82 engine swap — Open XML SDK-backed IWorkbook.
//
// This is THE v2.0.0 engine. It was grown additively alongside the NPOI-backed
// XssfWorkbook under the parallel-engine / late-cutover strategy (design I-82);
// since the cutover the default Workbook.Create() / Open() / OpenAsync()
// factories route here and the NPOI engine is deleted.
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
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlWorkbook : IWorkbook
{
    private readonly SpreadsheetDocument _document;
    private readonly MemoryStream _backing;
    private readonly WorkbookOptions _options;
    private readonly Dictionary<string, OoxmlSheet> _sheetsByName = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<OoxmlSheet> _sheetsByIndex = new();
    private OoxmlStylePool? _stylePool;
    private Dictionary<string, CellStyle>? _namedStyles;
    private bool? _date1904;
    private bool _disposed;
    private int _inMutation;
    // Relationship-orphan OPC parts captured at Open (decision #44 / SDK-quirk
    // #18): zip entries with a registered content type but no relationship
    // chain are invisible to the SDK's typed part graph and would be dropped
    // by the clone-based Save. Empty for created workbooks and for real-world
    // files (whose parts are rel-wired), so the common case carries no cost.
    private readonly List<(Uri Uri, string ContentType, byte[] Bytes)> _orphanParts = new();

    internal WorkbookOptions Options => _options;

    /// <summary>
    /// The workbook's style pool (xl/styles.xml), materialized on first use.
    /// On a created workbook this writes the Excel default scaffolding and the
    /// option-supplied default font; on an opened workbook it adopts the file's
    /// existing stylesheet (preserving its default font, lesson #8).
    /// </summary>
    internal OoxmlStylePool StylePool =>
        _stylePool ??= OoxmlStylePool.Attach(
            _document.WorkbookPart ?? throw new InvalidOperationException("Workbook has no workbook part."),
            _options);

    // Default number-format styles for date/time/duration cells (mirrors the
    // NPOI engine's DateStyle/DateTimeStyle/TimeStyle/DurationStyle, decisions
    // I-18/I-19, §7.9). Applied to an otherwise-unstyled cell by the date setters.
    internal static CellStyle DateStyleSpec { get; } = new() { NumberFormat = "yyyy-mm-dd" };
    internal static CellStyle DateTimeStyleSpec { get; } = new() { NumberFormat = "yyyy-mm-dd hh:mm:ss" };
    internal static CellStyle TimeStyleSpec { get; } = new() { NumberFormat = "h:mm:ss" };
    internal static CellStyle DurationStyleSpec { get; } = new() { NumberFormat = "[h]:mm:ss" };

    /// <summary>
    /// Whether the workbook uses the 1904 date system (workbookPr/@date1904).
    /// On an opened workbook the file is authoritative (lesson #9); on a created
    /// workbook it reflects <see cref="WorkbookOptions.DateSystem"/>.
    /// </summary>
    internal bool Date1904
    {
        get
        {
            if (_date1904 is { } cached) return cached;
            var pr = _document.WorkbookPart?.Workbook?.GetFirstChild<S.WorkbookProperties>();
            bool v = pr?.Date1904?.Value ?? false;
            _date1904 = v;
            return v;
        }
    }

    // ---- Excel date-serial conversion (1900/1904 epochs) --------------------
    // Matches the NPOI engine across the full Excel-representable range. The 1900
    // system reproduces Excel's fictitious 1900-02-29 leap day: dates on/after
    // 1900-03-01 carry the +1 phantom-day offset.

    private static readonly DateTime Epoch1900 = new(1899, 12, 31);
    private static readonly DateTime Epoch1904 = new(1904, 1, 1);
    private static readonly DateTime LeapBugThreshold = new(1900, 3, 1);

    internal double ToSerial(DateTime value) => ToSerial(value, Date1904);

    // Static form shared with the streaming engine (slice 9), whose date system
    // comes from StreamingOptions rather than a live workbookPr element.
    internal static double ToSerial(DateTime value, bool date1904)
    {
        if (date1904) return (value - Epoch1904).TotalDays;
        double serial = (value - Epoch1900).TotalDays;
        if (value >= LeapBugThreshold) serial += 1; // phantom 1900-02-29
        return serial;
    }

    internal DateTime FromSerial(double serial)
    {
        DateTime dt;
        if (Date1904)
        {
            dt = Epoch1904.AddDays(serial);
        }
        else
        {
            double s = serial >= 60 ? serial - 1 : serial; // remove phantom 1900-02-29
            dt = Epoch1900.AddDays(s);
        }
        return RoundToMillisecond(dt);
    }

    // The fraction-of-day serial (stored as a G17 double) accumulates sub-tick
    // error on the round-trip; round to the nearest millisecond so 13:45:00 does
    // not come back as 13:44:59.9999997. Mirrors NPOI's millisecond-resolution dates.
    private static DateTime RoundToMillisecond(DateTime dt)
    {
        long ticks = (dt.Ticks + (TimeSpan.TicksPerMillisecond / 2)) / TimeSpan.TicksPerMillisecond * TimeSpan.TicksPerMillisecond;
        return new DateTime(ticks, dt.Kind);
    }

    private OoxmlWorkbook(SpreadsheetDocument document, MemoryStream backing, WorkbookOptions options)
    {
        _document = document;
        _backing = backing;
        _options = options;
    }

    // ---- Factories (called by Workbook.Create / Open) ------------------------

    internal static OoxmlWorkbook Create(
        WorkbookOptions options,
        SpreadsheetDocumentType documentType = SpreadsheetDocumentType.Workbook)
    {
        ArgumentNullException.ThrowIfNull(options);
        var backing = new MemoryStream();
        try
        {
            // autoSave:false (R-35) — Save is explicit on this facade and the
            // backing stream is internal, so SDK autosave has nothing visible
            // to persist; left on, it re-serializes inside Dispose() and turns
            // a failed Save's exception into a second, unhandled throw at the
            // using-brace that masks the original.
            var document = SpreadsheetDocument.Create(backing, documentType, autoSave: false);
            var wbPart = document.AddWorkbookPart();
            wbPart.Workbook = new S.Workbook();
            // workbookPr (date system, lesson #9) must precede <sheets> in schema
            // order — insert it before appending the sheets container.
            if (options.DateSystem == DateSystem.Excel1904)
                wbPart.Workbook.AppendChild(new S.WorkbookProperties { Date1904 = true });
            wbPart.Workbook.AppendChild(new S.Sheets());

            var wb = new OoxmlWorkbook(document, backing, options);
            // Materialize the stylesheet now so font index 0 reflects the
            // workbook's default font/size (styles slice). On Open the pool adopts
            // the file's stylesheet instead, preserving its default font.
            _ = wb.StylePool;
            return wb;
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

        // Malformed-input gate, part 1 (SDK-quirk #17, the I-60 equivalent):
        // System.IO.Packaging opens an EMPTY read/write stream by CREATING a
        // brand-new package rather than throwing — explicitly reject it.
        if (backing.Length == 0)
        {
            backing.Dispose();
            throw new MalformedFileException(malformedMessage);
        }

        // Snapshot the source bytes before the SDK takes the backing stream —
        // the orphan-part capture below enumerates the package's raw OPC view,
        // which must not race the live document over the same stream.
        byte[] sourceBytes = backing.ToArray();

        SpreadsheetDocument document;
        try
        {
            // AutoSave=false (R-35): same rationale as Create — see the comment there.
            document = SpreadsheetDocument.Open(backing, isEditable: true, new OpenSettings { AutoSave = false });
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
        try
        {
            // Malformed-input gate, part 2 (SDK-quirk #17): a valid zip with no
            // workbook part (empty zip, random entries) opens without complaint —
            // require the workbook part up front. Package-level corruption (e.g.
            // a relationship pointing at a missing part) surfaces lazily as a raw
            // InvalidOperationException at first part access; the catch below
            // classifies it to MalformedFileException instead of leaking it.
            if (document.WorkbookPart is null)
                throw new MalformedFileException(malformedMessage);

            wb.IndexExistingSheets();
            wb.EnforceReadLimits();
            wb.CaptureOrphanParts(sourceBytes);
            return wb;
        }
        catch (WorkbookException)
        {
            // Our own contract exceptions (MalformedFileException above,
            // ResourceLimitExceededException from the read limits) propagate.
            wb.Dispose();
            throw;
        }
        catch (Exception ex) when (ex is InvalidOperationException || IsKnownMalformedOpenException(ex))
        {
            wb.Dispose();
            throw new MalformedFileException(malformedMessage, ex);
        }
        catch
        {
            wb.Dispose();
            throw;
        }
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
            // Resolve the worksheet part backing this sheet via its r:id —
            // fail-loud classification per the Open contract (R-37; the
            // pre-fix null-forgiving cast NRE'd on a corrupted r:id, found
            // by the deep-fuzz harness). A LEGAL chartsheet/dialogsheet
            // target also lands in the last branch until R-38 decides its
            // shape — honest rejection beats an InvalidCastException.
            var relId = sheet.Id?.Value;
            if (string.IsNullOrEmpty(relId))
                throw new MalformedFileException(
                    $"workbook sheet '{name}' is missing its relationship id (r:id)");
            OpenXmlPart part;
            try
            {
                part = wbPart.GetPartById(relId);
            }
            catch (Exception ex) when (ex is ArgumentException or KeyNotFoundException)
            {
                throw new MalformedFileException(
                    $"workbook sheet '{name}' references part '{relId}', which does not exist in the package", ex);
            }
            if (part is not WorksheetPart wsPart)
                throw new MalformedFileException(
                    $"workbook sheet '{name}' relationship '{relId}' targets a {part.GetType().Name}, not a worksheet part" +
                    " (chartsheet/dialogsheet workbooks are not supported yet — tracked as ledger R-38)");
            var wrapper = new OoxmlSheet(this, name, wsPart);
            // R-14: row/@r and c/@r are OPTIONAL per ECMA-376 (absent =
            // previous + 1). Our own writers always emit them, but a
            // spec-legal third-party file without them previously had its
            // rows/cells silently invisible to every reader path. Infer and
            // materialize the references once at open — Excel-compatible,
            // and every downstream consumer (row cache, dimension, sort,
            // EnumeratePopulated) sees one consistent grid.
            wrapper.NormalizeMissingReferences();
            _sheetsByIndex.Add(wrapper);
            _sheetsByName[name] = wrapper;
        }
    }

    /// <summary>
    /// Captures relationship-orphan OPC parts (decision #44 / SDK-quirk #18).
    /// The SDK part graph is relationship-defined, so a zip entry with a
    /// registered content type but no .rels chain — legal OPC, and pinned by
    /// the golden RoundTripPreservationTests — is invisible to
    /// <c>GetAllParts()</c> and silently dropped by the clone-based
    /// <see cref="Save(Stream, bool)"/>. This snapshots each orphan's URI,
    /// content type, and bytes from the source's raw packaging view;
    /// <see cref="ReinjectOrphanParts"/> restores them into every Save.
    /// </summary>
    private void CaptureOrphanParts(byte[] sourceBytes)
    {
        var reachable = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in _document.GetAllParts())
            reachable.Add(part.Uri.ToString());

        using var ms = new MemoryStream(sourceBytes, writable: false);
        using var pkg = System.IO.Packaging.Package.Open(ms, FileMode.Open, FileAccess.Read);
        foreach (var part in pkg.GetParts())
        {
            // Relationship parts are infrastructure, not content — the clone
            // re-emits the ones that matter and an orphan has none by definition.
            if (part.ContentType == "application/vnd.openxmlformats-package.relationships+xml")
                continue;
            if (reachable.Contains(part.Uri.ToString()))
                continue;

            using var s = part.GetStream(FileMode.Open, FileAccess.Read);
            using var buf = new MemoryStream();
            s.CopyTo(buf);
            _orphanParts.Add((part.Uri, part.ContentType, buf.ToArray()));
        }
    }

    /// <summary>
    /// Re-adds the captured orphan parts to the finalized clone bytes in
    /// <paramref name="tmp"/>. A part that has since become reachable (e.g.
    /// the consumer wired a relationship through the .Underlying escape
    /// hatch) is already in the clone and is skipped.
    /// </summary>
    private void ReinjectOrphanParts(MemoryStream tmp)
    {
        using var pkg = System.IO.Packaging.Package.Open(tmp, FileMode.Open, FileAccess.ReadWrite);
        foreach (var (uri, contentType, bytes) in _orphanParts)
        {
            if (pkg.PartExists(uri)) continue;
            var part = pkg.CreatePart(uri, contentType, System.IO.Packaging.CompressionOption.Normal);
            using var s = part.GetStream(FileMode.Create, FileAccess.Write);
            s.Write(bytes, 0, bytes.Length);
        }
    }

    /// <summary>
    /// WorkbookOptions read limits, enforced post-open (mirrors the NPOI
    /// engine's <c>XssfWorkbook.EnforceReadLimits</c>): <see cref="WorkbookOptions.ReadMaxSheets"/>
    /// and <see cref="WorkbookOptions.ReadMaxUncompressedBytes"/>. The byte check is
    /// best-effort over every relationship-reachable part — it catches over-the-line
    /// payloads after they've been buffered, the right place to fail loud given the
    /// whole package is already copied into the owned backing buffer.
    /// </summary>
    private void EnforceReadLimits()
    {
        // ReadMaxSheets
        int sheetCount = _sheetsByIndex.Count;
        if (sheetCount > _options.ReadMaxSheets)
        {
            throw new ResourceLimitExceededException(
                "sheet count", _options.ReadMaxSheets, sheetCount);
        }

        // ReadMaxUncompressedBytes
        long limit = _options.ReadMaxUncompressedBytes;
        if (limit <= 0) return;
        long total = 0;
        foreach (var part in _document.GetAllParts())
        {
            using var s = part.GetStream(FileMode.Open, FileAccess.Read);
            if (s.CanSeek)
            {
                total += s.Length;
            }
            else
            {
                // Length-by-read fallback for non-seekable part streams,
                // bounded so we never read past the configured limit.
                total += CountBytes(s, limit - total + 1);
            }

            if (total > limit)
            {
                throw new ResourceLimitExceededException(
                    "uncompressed package size in bytes", limit, total);
            }
        }
    }

    private static long CountBytes(Stream s, long capPlusOne)
    {
        byte[] buf = new byte[4096];
        long count = 0;
        int read;
        while ((read = s.Read(buf, 0, buf.Length)) > 0)
        {
            count += read;
            if (count >= capPlusOne) break;
        }
        return count;
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
        using var _ = EnterMutation();
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

        // Refresh every sheet's <dimension> from the live extent (R-13) —
        // sized output is part of the save contract on both engines.
        foreach (var sheet in _sheetsByIndex) sheet.UpdateDimension();

        // Flush the strongly-typed DOM into the in-memory package parts, then
        // clone into a throwaway buffer whose disposal finalizes the zip central
        // directory. The live document stays open and re-saveable.
        if (_document.CanSave) _document.Save();
        using (var tmp = new MemoryStream())
        {
            using (_document.Clone(tmp)) { }
            if (_orphanParts.Count > 0)
            {
                // The clone walked the relationship graph and dropped the
                // orphans (SDK-quirk #18) — restore them. Closing the package
                // may dispose tmp, so read via ToArray (valid on a disposed
                // MemoryStream) instead of repositioning.
                ReinjectOrphanParts(tmp);
                byte[] finished = tmp.ToArray();
                stream.Write(finished, 0, finished.Length);
            }
            else
            {
                tmp.Position = 0;
                tmp.CopyTo(stream);
            }
        }

        if (!leaveOpen) stream.Dispose();
    }

    public void Save(string path)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(path);
        // R-1: sibling-temp + rename — a failed serialization must never leave
        // the destination truncated or destroyed.
        AtomicFileWriter.Write(path, fs => Save(fs, leaveOpen: true));
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

    // Escape hatch (#32 / I-82): the live SDK document. Disposal is checked
    // BEFORE the value is returned (DisposedWorkbookMatrix contract).
    // Handing out the DOM invalidates the row caches (I-87) so mutations made
    // through the returned reference are observed by subsequent facade calls.
    public SpreadsheetDocument Underlying
    {
        get
        {
            ThrowIfDisposed();
            InvalidateRowCaches();
            return _document;
        }
    }

    // I-87: every escape-hatch getter routes here — any handed-out DOM node
    // can reach every sheet (and the stylesheet) via part traversal, so all
    // row caches and the style apply-memo reset.
    internal void InvalidateRowCaches()
    {
        foreach (var sheet in _sheetsByIndex) sheet.InvalidateRowCache();
        StylePool.ClearApplyMemo();
    }

    // Internal accessor for engine code that needs the document without the
    // public-hatch disposal ceremony (callers hold a live workbook).
    internal SpreadsheetDocument Document => _document;

    public void Dispose()
    {
        // R-15: in strict mode, disposal takes the same per-workbook lock as
        // every mutating path, so a concurrent mutation can never interleave
        // with teardown. The default mode stays opportunistic — its
        // documented race scope (WorkbookOptions.StrictConcurrencyDetection)
        // includes dispose-vs-mutate.
        if (Options.StrictConcurrencyDetection)
        {
            lock (_strictLock) { DisposeCore(); }
            return;
        }
        DisposeCore();
    }

    private void DisposeCore()
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

    /// <summary>
    /// Begins a mutating operation scope (mirrors the NPOI engine's
    /// <c>XssfWorkbook.EnterMutation</c>). Returns an <see cref="IDisposable"/>
    /// that releases the scope on dispose. If another thread is already inside
    /// a mutation when this is called, throws
    /// <see cref="InvalidOperationException"/> per decision #43.
    /// </summary>
    /// <remarks>
    /// The default mode is not a lock: nested same-thread mutations are also
    /// detected (the counter is process-global per workbook). Cell-level writes
    /// via <see cref="ICell"/> do not enter this scope; the counter guards
    /// higher-level structural mutations (AddSheet, AddNamedRange) where the
    /// read of internal collections — or, on this engine, the SDK part graph —
    /// during another thread's mutation would corrupt.
    /// </remarks>
    private readonly object _strictLock = new();

    internal MutationScope EnterMutation()
    {
        // Strict mode (decision I-59): take a real per-workbook lock.
        // Reentrant on the same thread (Monitor is reentrant), so nested
        // same-thread mutations are permitted in strict mode — unlike
        // the opportunistic counter which rejects them. The trade-off
        // is deliberate: strict mode is for "another thread cannot
        // silently corrupt me", not for "fluent-chain reentrancy is
        // forbidden". Callers concerned about reentrancy can compose
        // their own external guard.
        if (Options.StrictConcurrencyDetection)
        {
            Monitor.Enter(_strictLock);
            return new MutationScope(this, strict: true);
        }

        // Default opportunistic counter (decision #43).
        if (Interlocked.CompareExchange(ref _inMutation, 1, 0) != 0)
        {
            throw new InvalidOperationException(
                "Concurrent or reentrant mutation detected on IWorkbook. " +
                "Workbooks are not thread-safe (decision #43); serialize access externally. " +
                "Pass WorkbookOptions { StrictConcurrencyDetection = true } for a real-lock mode (decision I-59).");
        }
        return new MutationScope(this, strict: false);
    }

    internal readonly struct MutationScope : IDisposable
    {
        private readonly OoxmlWorkbook _owner;
        private readonly bool _strict;
        public MutationScope(OoxmlWorkbook owner, bool strict)
        {
            _owner = owner;
            _strict = strict;
        }
        public void Dispose()
        {
            if (_strict)
            {
                Monitor.Exit(_owner._strictLock);
            }
            else
            {
                Interlocked.Exchange(ref _owner._inMutation, 0);
            }
        }
    }

    // Named ranges (AddNamedRange / NamedRanges) land in OoxmlWorkbook.Names.cs
    // (I-82 structure slice).

    /// <summary>
    /// The display name of the sheet at <paramref name="zeroBasedIndex"/> in
    /// workbook (document) order — used by <see cref="OoxmlNamedRange.SheetScope"/>
    /// to resolve a <c>localSheetId</c>. Returns <c>null</c> if out of range.
    /// </summary>
    internal string? SheetNameAt(int zeroBasedIndex)
        => zeroBasedIndex >= 0 && zeroBasedIndex < _sheetsByIndex.Count
            ? _sheetsByIndex[zeroBasedIndex].Name
            : null;

    /// <summary>
    /// The workbook.xml <c>&lt;sheet&gt;</c> element backing <paramref name="part"/>,
    /// resolved via the part's relationship id. Sheet visibility (<c>state</c>) lives
    /// on this element, not on the worksheet — so <see cref="OoxmlSheet.Hidden"/>
    /// reaches it through here.
    /// </summary>
    internal S.Sheet SheetElementFor(WorksheetPart part)
    {
        var wbPart = _document.WorkbookPart
            ?? throw new InvalidOperationException("Workbook has no workbook part.");
        string rid = wbPart.GetIdOfPart(part);
        var sheets = wbPart.Workbook?.GetFirstChild<S.Sheets>()
            ?? throw new InvalidOperationException("Workbook has no <sheets> element.");
        foreach (var sheet in sheets.Elements<S.Sheet>())
            if (sheet.Id?.Value == rid) return sheet;
        throw new InvalidOperationException("No <sheet> element backs this worksheet part.");
    }

    // Resolves a sheet name to its 0-based document-order index (case-insensitive,
    // matching the workbook's sheet-name resolution), or -1 if no such sheet.
    private int IndexOfSheet(string name)
    {
        for (int i = 0; i < _sheetsByIndex.Count; i++)
            if (string.Equals(_sheetsByIndex[i].Name, name, StringComparison.OrdinalIgnoreCase))
                return i;
        return -1;
    }

    // ---- Style pool diagnostics + named-style registry (styles slice; OOXML
    // persistence in the closeout slice) -------------------------------------
    //
    // The in-process dictionary is the runtime source of truth (same model as
    // the NPOI engine's I-67): RegisterStyle records the name for ApplyNamedStyle
    // AND persists it into the stylesheet's cellStyleXfs/cellStyles tables (via
    // OoxmlStylePool.WriteNamedStyle) so names survive a save/open round-trip
    // and appear in Excel's Cell Styles panel. First access on an opened
    // workbook rehydrates the registry from the persisted entries.

    private Dictionary<string, CellStyle> NamedStyles
    {
        get
        {
            if (_namedStyles is null)
            {
                _namedStyles = new Dictionary<string, CellStyle>(StringComparer.OrdinalIgnoreCase);
                foreach (var entry in StylePool.ReadNamedStyles())
                    _namedStyles[entry.Key] = entry.Value;
            }
            return _namedStyles;
        }
    }

    public StylePoolDiagnostics GetStylePoolDiagnostics()
    {
        ThrowIfDisposed();
        var pool = StylePool;
        return new StylePoolDiagnostics(
            styleHits: pool.StyleHitCount,
            styleMisses: pool.StyleMissCount,
            fontHits: pool.FontHitCount,
            fontMisses: pool.FontMissCount,
            uniqueStyles: pool.UniqueStyles,
            uniqueFonts: pool.UniqueFonts);
    }

    public void RegisterStyle(string name, CellStyle style)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(style);
        if (name.Length == 0)
            throw new ArgumentException("Style name cannot be empty.", nameof(name));
        NamedStyles[name] = style;
        StylePool.WriteNamedStyle(name, style);
    }

    public CellStyle? GetRegisteredStyle(string name)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        return NamedStyles.TryGetValue(name, out var s) ? s : null;
    }

    public IReadOnlyCollection<string> RegisteredStyleNames
    {
        get { ThrowIfDisposed(); return NamedStyles.Keys; }
    }

    /// <summary>
    /// Resolves a registered style name to its <see cref="CellStyle"/>, throwing
    /// the canonical "no such name" error. Shared by ICell/IRange ApplyNamedStyle.
    /// </summary>
    internal CellStyle ResolveNamedStyleOrThrow(string name)
    {
        var style = GetRegisteredStyle(name);
        if (style is null)
            throw new ArgumentException(
                $"No style is registered under '{name}'. " +
                "Use IWorkbook.RegisterStyle before referencing the name.",
                nameof(name));
        return style;
    }

    // Protect / ProtectWithPassword / Unprotect / IsProtected / IsMacroEnabled
    // live in OoxmlWorkbook.Protection.cs (formulas/comments/hyperlinks slice).
    // SetThemeXml / GetThemeXml / ResolveThemeColor / GetThemeLineWidthEmu live in
    // OoxmlWorkbook.Theme.cs (drawings slice — theme round-trip).
}
