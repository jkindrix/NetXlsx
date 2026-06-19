// I-82 engine swap — Open XML SDK-backed IRange.
//
// Rectangular range with sparse (populated-only) default enumeration and dense
// EnumerateAll. Value(object?), ClearContents, Apply / ApplyNamedStyle, and
// Merge are all implemented. Bounds are normalized 1-based and inclusive.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace NetXlsx;

internal sealed class OoxmlRange : IRange
{
    private readonly OoxmlSheet _sheet;

    private readonly int _firstRow;
    private readonly int _lastRow;
    private readonly int _firstCol;
    private readonly int _lastCol;

    internal OoxmlRange(OoxmlSheet sheet, int row1, int col1, int row2, int col2)
    {
        _sheet = sheet;
        _firstRow = Math.Min(row1, row2);
        _lastRow = Math.Max(row1, row2);
        _firstCol = Math.Min(col1, col2);
        _lastCol = Math.Max(col1, col2);
    }

    private OoxmlWorkbook Wb => _sheet.WorkbookInternal;

    // Disposed-workbook guards per decision #42 — every IRange member,
    // including the bounds, throws after the owning workbook is disposed
    // (pinned by DisposedWorkbookMatrixTests on the NPOI engine).
    public int FirstRow { get { Wb.ThrowIfDisposed(); return _firstRow; } }
    public int LastRow { get { Wb.ThrowIfDisposed(); return _lastRow; } }
    public int FirstCol { get { Wb.ThrowIfDisposed(); return _firstCol; } }
    public int LastCol { get { Wb.ThrowIfDisposed(); return _lastCol; } }

    public string Address { get { Wb.ThrowIfDisposed(); return CellAddress.FormatRange(FirstRow, FirstCol, LastRow, LastCol); } }
    public int Count { get { Wb.ThrowIfDisposed(); return (LastRow - FirstRow + 1) * (LastCol - FirstCol + 1); } }
    public ISheet Sheet { get { Wb.ThrowIfDisposed(); return _sheet; } }

    public IEnumerator<ICell> GetEnumerator()
    {
        Wb.ThrowIfDisposed();
        return _sheet.EnumeratePopulated(FirstRow, FirstCol, LastRow, LastCol)
            .Cast<ICell>().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IEnumerable<ICell> EnumerateAll()
    {
        Wb.ThrowIfDisposed();
        return EnumerateAllCore();
    }

    private IEnumerable<ICell> EnumerateAllCore()
    {
        for (int r = FirstRow; r <= LastRow; r++)
            for (int c = FirstCol; c <= LastCol; c++)
                yield return _sheet.CellHandle(r, c);
    }

    public IRange Value(object? value)
    {
        Wb.ThrowIfDisposed();
        if (value is null) return ClearContents();
        foreach (var cell in EnumerateAllCore())
            ApplyValue(cell, value);
        return this;
    }

    /// <summary>
    /// Runtime-type-dispatched value assignment — mirrors the NPOI engine's
    /// <c>XssfRange.ApplyValue</c> scalar set and exception wording exactly
    /// (the "is not a supported scalar" message is a pinned v1 contract).
    /// </summary>
    private static void ApplyValue(ICell cell, object value)
    {
        switch (value)
        {
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

    public IRange ClearContents()
    {
        Wb.ThrowIfDisposed();
        // Materialize first: Clear() mutates the cells' DOM nodes during the walk.
        foreach (var cell in _sheet.EnumeratePopulated(FirstRow, FirstCol, LastRow, LastCol).ToList())
            cell.Clear();
        return this;
    }

    // ---- Styling (styles slice) --------------------------------------------

    public IRange Apply(CellStyle style)
    {
        Wb.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(style);
        // Dense: style every cell in the rectangle, materializing as needed —
        // matches the NPOI engine and is required for merged-cell border
        // rendering (lesson #4: borders render from boundary cells).
        foreach (var cell in EnumerateAllCore())
            cell.Style(style);
        return this;
    }

    public IRange ApplyNamedStyle(string name)
    {
        Wb.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        return Apply(_sheet.WorkbookInternal.ResolveNamedStyleOrThrow(name));
    }

    // ---- Merging ------------------------------------------------------------

    public IRange Merge()
    {
        Wb.ThrowIfDisposed();
        // Delegates to the sheet's merge surface (same overlap/1x1 contract;
        // NPOI-engine parity — XssfRange.Merge does exactly this).
        _sheet.MergeCells(Address);
        return this;
    }
}
