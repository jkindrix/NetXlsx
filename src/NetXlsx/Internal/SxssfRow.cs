using System;
using NPOI.XSSF.Streaming;

namespace NetXlsx;

internal sealed class SxssfRow : IStreamingRow
{
    private readonly SxssfWorkbook _workbook;
    private readonly SxssfSheet _sheet;
    private readonly SXSSFRow _underlying;
    private readonly int _row1;

    public SxssfRow(SxssfWorkbook workbook, SxssfSheet sheet, SXSSFRow underlying)
    {
        _workbook = workbook;
        _sheet = sheet;
        _underlying = underlying;
        _row1 = underlying.RowNum + 1;
    }

    public int Index { get { _workbook.ThrowIfDisposed(); return _row1; } }

    public IStreamingSheet Sheet { get { _workbook.ThrowIfDisposed(); return _sheet; } }

    public IStreamingCell Cell(int column)
    {
        _workbook.ThrowIfDisposed();
        if (column < 1 || column > CellAddress.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column,
                $"column must be in [1, {CellAddress.MaxColumn}]");
        return GetOrCreateCell(column);
    }

    public IStreamingCell this[int column] => Cell(column);

    public IStreamingCell this[string columnLetter]
    {
        get
        {
            _workbook.ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(columnLetter);
            int col = CellAddress.ParseColumn(columnLetter);
            return GetOrCreateCell(col);
        }
    }

    public IStreamingRow Set(int column, string value) { Cell(column).SetString(value); return this; }
    public IStreamingRow Set(int column, double value) { Cell(column).SetNumber(value); return this; }
    public IStreamingRow Set(int column, decimal value) { Cell(column).SetNumber(value); return this; }
    public IStreamingRow Set(int column, int value) { Cell(column).SetNumber(value); return this; }
    public IStreamingRow Set(int column, long value) { Cell(column).SetNumber(value); return this; }
    public IStreamingRow Set(int column, bool value) { Cell(column).SetBool(value); return this; }
    public IStreamingRow Set(int column, DateTime value) { Cell(column).SetDate(value); return this; }

    public void Flush()
    {
        _workbook.ThrowIfDisposed();
        // SXSSFSheet.FlushRows() flushes everything past the
        // row-access window. Public NPOI 2.7.x exposes only the
        // no-arg form; the (remaining, flushOnDisk) overload is
        // private.
        _sheet.Npoi.FlushRows();
    }

    private SxssfCell GetOrCreateCell(int col1)
    {
        int col0 = col1 - 1;
        var npoiCell = (SXSSFCell?)_underlying.GetCell(col0) ?? (SXSSFCell)_underlying.CreateCell(col0);
        return new SxssfCell(_workbook, npoiCell, _row1, col1);
    }
}
