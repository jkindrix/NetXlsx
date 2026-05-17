using System;
using NPOI.XSSF.Streaming;

namespace NetXlsx;

internal sealed class SxssfSheet : IStreamingSheet
{
    private readonly SxssfWorkbook _workbook;
    private readonly SXSSFSheet _underlying;
    private int _lastWritten0 = -1;   // 0-based; -1 == no rows yet

    public SxssfSheet(SxssfWorkbook workbook, SXSSFSheet underlying)
    {
        _workbook = workbook;
        _underlying = underlying;
    }

    public string Name { get { _workbook.ThrowIfDisposed(); return _underlying.SheetName; } }

    public IStreamingWorkbook Workbook { get { _workbook.ThrowIfDisposed(); return _workbook; } }

    public IStreamingRow AppendRow()
    {
        _workbook.ThrowIfDisposed();
        int next0 = _lastWritten0 + 1;
        return CreateAt0(next0);
    }

    public IStreamingRow AppendRow(int index)
    {
        _workbook.ThrowIfDisposed();
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

    private SxssfRow CreateAt0(int row0)
    {
        if (row0 + 1 > CellAddress.MaxRow)
            throw new InvalidOperationException(
                $"appending would exceed Excel's row limit of {CellAddress.MaxRow}");
        var npoiRow = (SXSSFRow)_underlying.CreateRow(row0);
        _lastWritten0 = row0;
        return new SxssfRow(_workbook, this, npoiRow);
    }

    public SXSSFSheet Underlying { get { _workbook.ThrowIfDisposed(); return _underlying; } }
}
