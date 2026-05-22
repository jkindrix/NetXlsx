// Internal wrapper over NPOI.XSSF.UserModel.XSSFWorkbook.
// One implementation of IWorkbook for v0.2.0; v0.3.0 may add a streaming
// counterpart (IStreamingWorkbook / SXSSF) per design decision #7.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed partial class XssfWorkbook : IWorkbook
{
    private readonly XSSFWorkbook _underlying;
    private readonly Dictionary<string, XssfSheet> _sheetsByName = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<XssfSheet> _sheetsByIndex = new();
    private bool _disposed;
    // Decision #43: workbooks are not thread-safe, but we detect concurrent
    // mutation via a non-locking reentry counter. A thread entering a
    // mutating operation increments _inMutation atomically; if another
    // thread sees a non-zero value when it tries to mutate, it throws
    // InvalidOperationException instead of silently corrupting state.
    private int _inMutation;

    // Style-pool dedup (decision #4 / spike 1). All ICellStyle
    // allocations — including the date/time defaults (decisions I-18,
    // I-19, §7.9) — flow through this single pool, replacing the S29
    // interim cache.
    private CellStylePool? _stylePool;
    internal CellStylePool StylePool => _stylePool ??= new CellStylePool(_underlying);

    private readonly WorkbookOptions _options;
    internal WorkbookOptions Options => _options;

    public XssfWorkbook(XSSFWorkbook underlying) : this(underlying, new WorkbookOptions()) { }

    public XssfWorkbook(XSSFWorkbook underlying, WorkbookOptions options)
    {
        ArgumentNullException.ThrowIfNull(underlying);
        ArgumentNullException.ThrowIfNull(options);
        _underlying = underlying;
        _options = options;

        // Apply DefaultFontName / DefaultFontSize to the workbook's
        // default font (index 0). NPOI auto-creates a default font on
        // first XSSFWorkbook construction; we mutate it in place.
        if (_underlying.NumberOfFonts > 0)
        {
            var defaultFont = _underlying.GetFontAt(0);
            defaultFont.FontName = _options.DefaultFontName;
            defaultFont.FontHeightInPoints = (short)_options.DefaultFontSize;
        }

        // Read-side safety checks (zip-bomb defense + sheet-count cap).
        // These run on the Open path (NumberOfSheets > 0 means the
        // underlying workbook was constructed from an existing file).
        if (_underlying.NumberOfSheets > 0)
        {
            EnforceReadLimits();
        }

        // Index any sheets that already exist (the Open path).
        for (int i = 0; i < _underlying.NumberOfSheets; i++)
        {
            var npoiSheet = (XSSFSheet)_underlying.GetSheetAt(i);
            var wrapper = new XssfSheet(this, npoiSheet);
            _sheetsByIndex.Add(wrapper);
            _sheetsByName[npoiSheet.SheetName] = wrapper;
        }
    }

    private void EnforceReadLimits()
    {
        // ReadMaxSheets
        int sheetCount = _underlying.NumberOfSheets;
        if (sheetCount > _options.ReadMaxSheets)
        {
            throw new ResourceLimitExceededException(
                "sheet count", _options.ReadMaxSheets, sheetCount);
        }

        // ReadMaxUncompressedBytes — best-effort post-Open check.
        // NPOI doesn't expose total uncompressed size directly; we sum
        // each OPC part's stream length. This catches over-the-line
        // payloads after they've been buffered in memory, which is the
        // right place to fail loud given that NPOI has no streaming
        // open API. Pre-buffer zip-bomb defense would need OPC-level
        // inspection before NPOI parses — out of scope for v1.
        long limit = _options.ReadMaxUncompressedBytes;
        if (limit <= 0) return;
        long total = 0;
        foreach (var part in _underlying.Package.GetParts())
        {
            System.IO.Stream? s = null;
            try
            {
                s = part.GetInputStream();
            }
            catch (InvalidOperationException)
            {
                // NPOI's PackagePropertiesPart (core/extended/custom
                // properties) doesn't expose a generic input stream —
                // GetInputStreamImpl throws "Operation not authorized."
                // These parts are bounded-small (core.xml etc.); skip
                // them rather than fail the open.
                continue;
            }

            try
            {
                if (s.CanSeek)
                {
                    total += s.Length;
                }
                else
                {
                    // Fall back to length-by-CopyTo when the stream is
                    // not seekable. Bounded so we never read more than
                    // we'd accept.
                    total += CountBytes(s, limit - total + 1);
                }
            }
            finally
            {
                s.Dispose();
            }

            if (total > limit)
            {
                throw new ResourceLimitExceededException(
                    "uncompressed package size in bytes", limit, total);
            }
        }
    }

    private static long CountBytes(System.IO.Stream s, long capPlusOne)
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

        var npoiSheet = (XSSFSheet)_underlying.CreateSheet(name);
        var wrapper = new XssfSheet(this, npoiSheet);
        _sheetsByIndex.Add(wrapper);
        _sheetsByName[name] = wrapper;
        return wrapper;
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

        // NPOI's XSSFWorkbook.Write closes the stream by default (a Java
        // POI carry-over). Use the bool overload to leave it open; we
        // restore the caller's intent below.
        _underlying.Write(stream, leaveOpen: true);
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

    public INamedRange AddNamedRange(string name, string formula, string? sheetScope = null)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(formula);
        if (name.Length == 0)
            throw new ArgumentException("name cannot be empty", nameof(name));
        if (formula.Length == 0)
            throw new ArgumentException("formula cannot be empty", nameof(formula));

        using var _ = EnterMutation();

        int? sheetIndex = null;
        if (sheetScope is not null)
        {
            int idx = _underlying.GetSheetIndex(sheetScope);
            if (idx < 0)
                throw new SheetNameException(sheetScope, "no sheet with that name exists in this workbook (sheetScope)");
            sheetIndex = idx;
        }

        // Excel itself permits a workbook-scope name and a same-text
        // sheet-scope name to coexist, but NPOI 2.7.x rejects this at
        // XSSFName.ValidateName ("The workbook already contains this
        // name"). v1 therefore requires names to be unique workbook-wide
        // regardless of scope. Documented in implementation-notes.md.
        foreach (var existing in _underlying.GetAllNames())
        {
            if (string.Equals(existing.NameName, name, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"a named range '{name}' already exists in the workbook " +
                    "(case-insensitive). NPOI 2.7.x requires names to be unique " +
                    "workbook-wide regardless of scope.", nameof(name));
        }

        // Strip an optional leading '=' for consistency with SetFormula.
        var body = formula.Length > 0 && formula[0] == '=' ? formula.Substring(1) : formula;

        NPOI.SS.UserModel.IName npoiName;
        try
        {
            npoiName = _underlying.CreateName();
            npoiName.NameName = name;
            if (sheetIndex is int si) npoiName.SheetIndex = si;
            npoiName.RefersToFormula = body;
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"invalid named range '{name}' = '{formula}': {ex.Message}", nameof(name), ex);
        }

        return new XssfNamedRange(this, npoiName);
    }

    public IReadOnlyList<INamedRange> NamedRanges
    {
        get
        {
            ThrowIfDisposed();
            var all = _underlying.GetAllNames();
            if (all.Count == 0) return Array.Empty<INamedRange>();
            var list = new INamedRange[all.Count];
            for (int i = 0; i < all.Count; i++)
                list[i] = new XssfNamedRange(this, all[i]);
            return list;
        }
    }

    public XSSFWorkbook Underlying
    {
        get { ThrowIfDisposed(); return _underlying; }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _underlying.Close();
    }

    internal void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(IWorkbook));
    }

    /// <summary>
    /// Begins a mutating operation scope. Returns an <see cref="IDisposable"/>
    /// that decrements the reentry counter on dispose. If another thread is
    /// already inside a mutation when this is called, throws
    /// <see cref="InvalidOperationException"/> per decision #43.
    /// </summary>
    /// <remarks>
    /// This is not a lock: nested same-thread mutations are also detected
    /// (the counter is process-global per workbook). That's a feature —
    /// nested mutations from a single thread (e.g., a callback within a
    /// fluent chain that mutates the workbook recursively) are equally
    /// undefined and rejected. Cell-level writes via <see cref="ICell"/>
    /// do not enter this scope; the counter guards higher-level structural
    /// mutations (AddSheet, future Add/RemoveRow, etc.) where the read of
    /// internal collections during another thread's mutation would corrupt.
    /// </remarks>
    internal MutationScope EnterMutation()
    {
        if (Interlocked.CompareExchange(ref _inMutation, 1, 0) != 0)
        {
            throw new InvalidOperationException(
                "Concurrent or reentrant mutation detected on IWorkbook. " +
                "Workbooks are not thread-safe (decision #43); serialize access externally.");
        }
        return new MutationScope(this);
    }

    internal readonly struct MutationScope : IDisposable
    {
        private readonly XssfWorkbook _owner;
        public MutationScope(XssfWorkbook owner) { _owner = owner; }
        public void Dispose() => Interlocked.Exchange(ref _owner._inMutation, 0);
    }

    // Date/time default styles are now pool entries (S29 → folded into
    // the full pool with the styling slice).
    internal ICellStyle DateStyle => StylePool.GetOrCreate(new CellStyle { NumberFormat = "yyyy-mm-dd" });
    internal ICellStyle DateTimeStyle => StylePool.GetOrCreate(new CellStyle { NumberFormat = "yyyy-mm-dd hh:mm:ss" });
    internal ICellStyle TimeStyle => StylePool.GetOrCreate(new CellStyle { NumberFormat = "h:mm:ss" });
    internal ICellStyle DurationStyle => StylePool.GetOrCreate(new CellStyle { NumberFormat = "[h]:mm:ss" });
}
