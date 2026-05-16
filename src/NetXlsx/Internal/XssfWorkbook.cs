// Internal wrapper over NPOI.XSSF.UserModel.XSSFWorkbook.
// One implementation of IWorkbook for v0.2.0; v0.3.0 may add a streaming
// counterpart (IStreamingWorkbook / SXSSF) per design decision #7.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfWorkbook : IWorkbook
{
    private readonly XSSFWorkbook _underlying;
    private readonly Dictionary<string, XssfSheet> _sheetsByName = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<XssfSheet> _sheetsByIndex = new();
    private bool _disposed;

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
}
