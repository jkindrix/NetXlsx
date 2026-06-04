using System;
using System.Collections;
using System.Collections.Generic;

namespace NetXlsx;

/// <summary>
/// Rectangular range of cells. Backed by the sheet; lazy on iteration.
/// Whole-row / whole-column ranges materialize lazily — iterating them
/// densely is 1M+ items by design.
/// </summary>
internal sealed class XssfRange : IRange
{
    private readonly XssfWorkbook _workbook;
    private readonly XssfSheet _sheet;
    private readonly int _row1;
    private readonly int _col1;
    private readonly int _row2;
    private readonly int _col2;

    public XssfRange(XssfWorkbook workbook, XssfSheet sheet, int row1, int col1, int row2, int col2)
    {
        _workbook = workbook;
        _sheet = sheet;
        _row1 = row1;
        _col1 = col1;
        _row2 = row2;
        _col2 = col2;
    }

    public string Address
    {
        get { _workbook.ThrowIfDisposed(); return CellAddress.FormatRange(_row1, _col1, _row2, _col2); }
    }

    public int FirstRow { get { _workbook.ThrowIfDisposed(); return _row1; } }
    public int LastRow  { get { _workbook.ThrowIfDisposed(); return _row2; } }
    public int FirstCol { get { _workbook.ThrowIfDisposed(); return _col1; } }
    public int LastCol  { get { _workbook.ThrowIfDisposed(); return _col2; } }

    public int Count
    {
        get
        {
            _workbook.ThrowIfDisposed();
            // Dense count of coordinates in the rectangle.
            return (_row2 - _row1 + 1) * (_col2 - _col1 + 1);
        }
    }

    public ISheet Sheet { get { _workbook.ThrowIfDisposed(); return _sheet; } }

    /// <summary>
    /// Sparse iteration — yields cells that the underlying NPOI sheet
    /// has actually materialized within the rectangle. Used by the base
    /// <see cref="IEnumerable{ICell}"/> implementation.
    /// </summary>
    public IEnumerator<ICell> GetEnumerator()
    {
        _workbook.ThrowIfDisposed();
        var npoiSheet = _sheet.Npoi;
        int rowEnd0 = Math.Min(_row2 - 1, npoiSheet.LastRowNum);

        for (int r0 = _row1 - 1; r0 <= rowEnd0; r0++)
        {
            var npoiRow = npoiSheet.GetRow(r0);
            if (npoiRow is null) continue;

            // NPOI's IRow.FirstCellNum / LastCellNum are short-typed.
            int firstCol0 = Math.Max(_col1 - 1, npoiRow.FirstCellNum);
            int lastCol0 = Math.Min(_col2 - 1, npoiRow.LastCellNum - 1);

            for (int c0 = firstCol0; c0 <= lastCol0; c0++)
            {
                var npoiCell = npoiRow.GetCell(c0);
                if (npoiCell is null) continue;
                yield return _sheet[r0 + 1, c0 + 1];
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IEnumerable<ICell> EnumerateAll()
    {
        _workbook.ThrowIfDisposed();
        for (int r = _row1; r <= _row2; r++)
        {
            for (int c = _col1; c <= _col2; c++)
            {
                yield return _sheet[r, c];
            }
        }
    }

    public IRange Value(object? value)
    {
        _workbook.ThrowIfDisposed();
        foreach (var cell in EnumerateAll())
        {
            ApplyValue(cell, value);
        }
        return this;
    }

    /// <summary>
    /// Runtime-type-dispatched value assignment. Mirrors the
    /// <see cref="IRow.Set"/> overload set with the addition of <c>null</c>
    /// (clears the cell). Unsupported types throw with a message that
    /// names the offending type — callers should reach for the typed
    /// setters on <see cref="ICell"/> directly.
    /// </summary>
    private static void ApplyValue(ICell cell, object? value)
    {
        switch (value)
        {
            case null:           cell.Clear(); break;
            case string s:       cell.SetString(s); break;
            case bool b:         cell.SetBool(b); break;
            case int i:          cell.SetNumber(i); break;
            case long l:         cell.SetNumber(l); break;
            case short sh:       cell.SetNumber((int)sh); break;
            case byte by:        cell.SetNumber((int)by); break;
            case sbyte sb:       cell.SetNumber((int)sb); break;
            case ushort us:      cell.SetNumber((int)us); break;
            case uint ui:        cell.SetNumber((long)ui); break;
            case ulong ul:       cell.SetNumber((long)ul); break;
            case double d:       cell.SetNumber(d); break;
            case float f:        cell.SetNumber((double)f); break;
            case decimal m:      cell.SetNumber(m); break;
            case DateTime dt:    cell.SetDate(dt); break;
            case DateOnly d1:    cell.SetDate(d1); break;
            case TimeOnly t:     cell.SetTime(t); break;
            case TimeSpan ts:    cell.SetDuration(ts); break;
            default:
                throw new ArgumentException(
                    $"IRange.Value: type '{value.GetType()}' is not a supported scalar. " +
                    $"Use the typed setters on ICell for custom types, or pass null to clear.",
                    nameof(value));
        }
    }

    public IRange Apply(CellStyle style)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(style);
        foreach (var cell in EnumerateAll())
        {
            cell.Style(style);
        }
        return this;
    }

    public IRange ApplyNamedStyle(string name)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        return Apply(_workbook.ResolveNamedStyleOrThrow(name));
    }

    public IRange Merge()
    {
        _workbook.ThrowIfDisposed();
        _sheet.MergeCells(Address);
        return this;
    }

    public IRange ClearContents()
    {
        _workbook.ThrowIfDisposed();
        // Sparse — clearing already-empty cells is wasted work and would
        // materialize backing storage we'd rather leave absent.
        foreach (var cell in this)
        {
            cell.Clear();
        }
        return this;
    }
}
