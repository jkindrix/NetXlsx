// I-82 engine swap — Open XML SDK-backed IRow.
//
// A row handle over a 1-based row index. Cell access and value setters delegate
// to OoxmlCell; HeightInPoints and Hidden write the SDK <row> attributes
// (ht/customHeight, hidden) directly. The DateTime/DateOnly/TimeOnly/TimeSpan
// Set overloads delegate to the cell's date/time setters.

using System;

namespace NetXlsx;

internal sealed class OoxmlRow : IRow
{
    private readonly OoxmlSheet _sheet;
    private readonly int _index;

    internal OoxmlRow(OoxmlSheet sheet, int index)
    {
        _sheet = sheet;
        _index = index;
    }

    public int Index { get { _sheet.WorkbookInternal.ThrowIfDisposed(); return _index; } }
    public ISheet Sheet { get { _sheet.WorkbookInternal.ThrowIfDisposed(); return _sheet; } }

    public ICell Cell(int column)
    {
        _sheet.WorkbookInternal.ThrowIfDisposed();
        // Effective cap: min(user-configured option, Excel hard cap) — mirrors XssfRow.
        int colCap = Math.Min(_sheet.WorkbookInternal.Options.MaxColsPerSheet, CellAddress.MaxColumn);
        if (column < 1 || column > colCap)
            throw new ArgumentOutOfRangeException(nameof(column), column, $"column must be in [1, {colCap}]");
        return new OoxmlCell(_sheet, _index, column);
    }

    public ICell this[int column] => Cell(column);
    public ICell this[string columnLetter] => Cell(CellAddress.ParseColumn(columnLetter));

    public IRow Set(int column, string value) { Cell(column).SetString(value); return this; }
    public IRow Set(int column, double value) { Cell(column).SetNumber(value); return this; }
    public IRow Set(int column, decimal value) { Cell(column).SetNumber(value); return this; }
    public IRow Set(int column, int value) { Cell(column).SetNumber(value); return this; }
    public IRow Set(int column, long value) { Cell(column).SetNumber(value); return this; }
    public IRow Set(int column, bool value) { Cell(column).SetBool(value); return this; }
    public IRow Set(int column, DateTime value) { Cell(column).SetDate(value); return this; }
    public IRow Set(int column, DateOnly value) { Cell(column).SetDate(value); return this; }
    public IRow Set(int column, TimeOnly value) { Cell(column).SetTime(value); return this; }
    public IRow Set(int column, TimeSpan value) { Cell(column).SetDuration(value); return this; }

    // ---- Row layout (styles slice) -----------------------------------------

    // Excel's default row height in points when a row carries no explicit height.
    private const float DefaultHeightPoints = 15f;

    private OoxmlWorkbook Wb => _sheet.WorkbookInternal;

    public float HeightInPoints
    {
        get
        {
            Wb.ThrowIfDisposed();
            return (float)(_sheet.FindRow(_index)?.Height?.Value ?? DefaultHeightPoints);
        }
        set
        {
            Wb.ThrowIfDisposed();
            var row = _sheet.GetOrCreateRow(_index);
            row.Height = value;
            // customHeight is the user-pinned signal (lesson #7) — set it so the
            // height survives round-trip regardless of value.
            row.CustomHeight = true;
        }
    }

    public bool Hidden
    {
        get { Wb.ThrowIfDisposed(); return _sheet.FindRow(_index)?.Hidden?.Value ?? false; }
        set
        {
            Wb.ThrowIfDisposed();
            _sheet.GetOrCreateRow(_index).Hidden = value ? true : null;
        }
    }

    // Escape hatch (#32 / I-82): the row element. Reaching for the raw node
    // is a write-like act — a never-materialized row materializes here.
    // Materialize first, then invalidate the row caches (I-87) so mutations
    // made through the returned reference are observed by the facade.
    public DocumentFormat.OpenXml.Spreadsheet.Row Underlying
    {
        get
        {
            _sheet.WorkbookInternal.ThrowIfDisposed();
            var row = _sheet.GetOrCreateRow(_index);
            _sheet.WorkbookInternal.InvalidateRowCaches();
            return row;
        }
    }
}
