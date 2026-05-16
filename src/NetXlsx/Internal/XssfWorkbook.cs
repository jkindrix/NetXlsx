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

internal sealed class XssfWorkbook : IWorkbook
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

    public XssfWorkbook(XSSFWorkbook underlying)
    {
        ArgumentNullException.ThrowIfNull(underlying);
        _underlying = underlying;
        // Index any sheets that already exist (the Open path).
        for (int i = 0; i < _underlying.NumberOfSheets; i++)
        {
            var npoiSheet = (XSSFSheet)_underlying.GetSheetAt(i);
            var wrapper = new XssfSheet(this, npoiSheet);
            _sheetsByIndex.Add(wrapper);
            _sheetsByName[npoiSheet.SheetName] = wrapper;
        }
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
