using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class SxssfWorkbook : IStreamingWorkbook
{
    private readonly SXSSFWorkbook _underlying;
    private readonly XSSFWorkbook _xssfBase;
    private readonly StreamingOptions _options;
    private readonly Dictionary<string, SxssfSheet> _sheetsByName = new(StringComparer.OrdinalIgnoreCase);
    private CellStylePool? _stylePool;
    private bool _disposed;

    public SxssfWorkbook(StreamingOptions options)
    {
        _options = options;
        _xssfBase = new XSSFWorkbook();
        // Streaming workbooks are always newly created (write-only), so the
        // 1904 epoch from options applies unconditionally (design #15).
        XssfWorkbook.ApplyDateSystem(_xssfBase, options.DateSystem);
        _underlying = new SXSSFWorkbook(_xssfBase, options.RowAccessWindowSize)
        {
            CompressTempFiles = options.CompressTempFiles,
        };
    }

    internal XSSFWorkbook XssfBase => _xssfBase;
    internal CellStylePool StylePool => _stylePool ??= new CellStylePool(_xssfBase);
    internal StreamingOptions Options => _options;

    public IStreamingSheet AddSheet(string name)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        Workbook.ValidateSheetName(name);
        if (_sheetsByName.ContainsKey(name))
            throw new SheetNameException(name, "a sheet with this name already exists (case-insensitive)");

        var npoiSheet = (SXSSFSheet)_underlying.CreateSheet(name);
        var wrapper = new SxssfSheet(this, npoiSheet);
        _sheetsByName[name] = wrapper;
        return wrapper;
    }

    public void Save(string path)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(path);
        using var fs = File.Create(path);
        Save(fs, leaveOpen: false);
    }

    public void Save(Stream stream, bool leaveOpen = true)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));

        // SXSSF's Write closes the stream by default — same Java-POI
        // carry-over as XSSF.
        _underlying.Write(stream, leaveOpen: true);
        if (!leaveOpen) stream.Dispose();
    }

    public Task SaveAsync(string path, CancellationToken ct = default)
    {
        ThrowIfDisposed();
        ct.ThrowIfCancellationRequested();
        return Task.Run(() => Save(path), ct);
    }

    public Task SaveAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default)
    {
        ThrowIfDisposed();
        ct.ThrowIfCancellationRequested();
        return Task.Run(() => Save(stream, leaveOpen), ct);
    }

    public SXSSFWorkbook Underlying
    {
        get { ThrowIfDisposed(); return _underlying; }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        // Ordering is load-bearing: SXSSFWorkbook.Dispose() deletes the
        // streaming temp files and MUST run before Close(). Calling Close()
        // first makes the subsequent Dispose() throw ObjectDisposedException
        // ("Cannot write to a closed TextWriter") in NPOI 2.7.3. Each step
        // is best-effort via finally so a throw in one still releases the
        // rest (temp files / package handles) rather than leaking them.
        try
        {
            _underlying.Dispose();
        }
        finally
        {
            try { _underlying.Close(); }
            finally { _xssfBase.Close(); }
        }
    }

    public ValueTask DisposeAsync()
    {
        Dispose();
        return ValueTask.CompletedTask;
    }

    internal void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }
}
