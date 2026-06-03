// I-82 engine swap — streaming sheet (slice 9). Owns the per-sheet temp file
// and the forward-only OpenXmlPartWriter over it, plus the bounded row-access
// window. See OoxmlStreamingWorkbook for the architecture note.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlStreamingSheet : IStreamingSheet, IDisposable
{
    private readonly OoxmlStreamingWorkbook _workbook;
    private readonly string _name;
    private readonly string _tempPath;
    private readonly int _windowSize;   // <= 0 means unbounded (NPOI -1 semantics)

    // Rows not yet flushed to the temp file, in append (= ascending row) order.
    private readonly Queue<StreamingRowBuffer> _window = new();

    private Stream? _tempStream;
    private OpenXmlWriter? _writer;
    private int _lastWritten0 = -1;     // 0-based; -1 == no rows yet
    private bool _finalized;

    internal OoxmlStreamingSheet(OoxmlStreamingWorkbook workbook, string name)
    {
        _workbook = workbook;
        _name = name;
        _windowSize = workbook.Options.RowAccessWindowSize;
        _tempPath = Path.Combine(Path.GetTempPath(), $"netxlsx-stream-{Guid.NewGuid():N}.xml.tmp");

        // The worksheet prolog is written immediately; rows follow as they
        // flush out of the window. CompressTempFiles trades CPU for disk I/O,
        // mirroring SXSSF's gzip temp files.
        Stream fs = new FileStream(_tempPath, FileMode.Create, FileAccess.Write, FileShare.None);
        if (workbook.Options.CompressTempFiles)
            fs = new GZipStream(fs, CompressionLevel.Fastest);
        _tempStream = fs;
        _writer = new OpenXmlPartWriter(fs);
        _writer.WriteStartDocument();
        _writer.WriteStartElement(new S.Worksheet());
        _writer.WriteStartElement(new S.SheetData());
    }

    internal string SheetName => _name;

    public string Name
    {
        get { _workbook.ThrowIfDisposed(); return _name; }
    }

    public IStreamingWorkbook Workbook
    {
        get { _workbook.ThrowIfDisposed(); return _workbook; }
    }

    public IStreamingRow AppendRow()
    {
        _workbook.ThrowIfDisposed();
        _workbook.ThrowIfSaved();
        return CreateAt0(_lastWritten0 + 1);
    }

    public IStreamingRow AppendRow(int index)
    {
        _workbook.ThrowIfDisposed();
        _workbook.ThrowIfSaved();
        if (index < 1 || index > CellAddress.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(index), index,
                $"row index must be in [1, {CellAddress.MaxRow}]");
        int next0 = index - 1;
        if (next0 <= _lastWritten0)
            throw new ArgumentOutOfRangeException(nameof(index), index,
                $"row {index} cannot be revisited — last written row was {_lastWritten0 + 1}. " +
                "Streaming rows are append-only (decision #7 / design §6.3).");
        return CreateAt0(next0);
    }

    public NPOI.XSSF.Streaming.SXSSFSheet Underlying
    {
        get
        {
            _workbook.ThrowIfDisposed();
            throw new NotSupportedException(
                "IStreamingSheet.Underlying (NPOI SXSSFSheet) is not available on the " +
                "Open XML SDK engine (I-82). The streaming engine writes forward-only " +
                "through OpenXmlWriter; there is no live sheet object to expose.");
        }
    }

    private OoxmlStreamingRow CreateAt0(int row0)
    {
        if (row0 + 1 > CellAddress.MaxRow)
            throw new InvalidOperationException(
                $"appending would exceed Excel's row limit of {CellAddress.MaxRow}");

        // Bounded window: flushing the oldest row(s) keeps at most
        // RowAccessWindowSize rows in memory — the same knob SXSSF honors.
        if (_windowSize > 0)
        {
            while (_window.Count >= _windowSize)
                FlushOldest();
        }

        var buffer = new StreamingRowBuffer(row0 + 1);
        _window.Enqueue(buffer);
        _lastWritten0 = row0;
        return new OoxmlStreamingRow(_workbook, this, buffer);
    }

    // ---- Flushing --------------------------------------------------------------

    /// <summary>
    /// Flushes every buffered row to the temp file (IStreamingRow.Flush /
    /// SXSSFSheet.FlushRows() semantics). Flushed rows reject further writes —
    /// forward-only is enforced, not faked (design.md I-82, streaming slice).
    /// </summary>
    internal void FlushAllBuffered()
    {
        if (_finalized) return; // post-Save: everything is already on disk
        while (_window.Count > 0)
            FlushOldest();
    }

    private void FlushOldest()
    {
        var buffer = _window.Dequeue();
        _writer!.WriteElement(Materialize(buffer));
        buffer.MarkFlushed(); // releases the cell data; later writes fail loud
    }

    private static S.Row Materialize(StreamingRowBuffer buffer)
    {
        var row = new S.Row { RowIndex = (uint)buffer.Row1 };
        foreach (var (col, cell) in buffer.Cells) // SortedDictionary: ascending columns
            row.AppendChild(cell.ToCell(CellAddress.Format(buffer.Row1, col)));
        return row;
    }

    // ---- Save-time finalization / cleanup ---------------------------------------

    /// <summary>
    /// Flushes the remaining window, closes the worksheet XML, and releases the
    /// writer. Idempotent. After this the temp file holds the complete part.
    /// </summary>
    internal void FinalizeXml()
    {
        if (_finalized) return;
        while (_window.Count > 0)
            FlushOldest();
        _writer!.WriteEndElement(); // </sheetData>
        _writer.WriteEndElement();  // </worksheet>
        _writer.Close();
        _writer = null;
        _tempStream!.Dispose();     // flushes the gzip trailer when compressing
        _tempStream = null;
        _finalized = true;
    }

    /// <summary>Opens the finalized worksheet XML for FeedData assembly.</summary>
    internal Stream OpenFinalizedXml()
    {
        Stream fs = File.OpenRead(_tempPath);
        return _workbook.Options.CompressTempFiles
            ? new GZipStream(fs, CompressionMode.Decompress)
            : fs;
    }

    /// <summary>The workbook owns sheet lifetime; its Dispose fans out here.</summary>
    public void Dispose() => DisposeTemp();

    /// <summary>Best-effort writer shutdown + temp-file deletion (workbook Dispose).</summary>
    internal void DisposeTemp()
    {
        try { _writer?.Close(); }
        catch (Exception) { /* best-effort: a failed flush must not block cleanup */ }
        _writer = null;
        try { _tempStream?.Dispose(); }
        catch (Exception) { /* best-effort */ }
        _tempStream = null;
        try { File.Delete(_tempPath); }
        catch (Exception) { /* best-effort: orphaned temp files are the OS's to reap */ }
    }
}
