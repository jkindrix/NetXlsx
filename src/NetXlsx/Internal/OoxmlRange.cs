// I-82 engine swap — Open XML SDK-backed IRange (cells & rows slice).
//
// Rectangular range with sparse (populated-only) default enumeration and dense
// EnumerateAll. Value(object?) and ClearContents are cell-value operations and
// are implemented; Apply / ApplyNamedStyle (styles slice) and Merge (merge
// slice) throw NotYet. Bounds are normalized 1-based and inclusive.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace NetXlsx;

internal sealed class OoxmlRange : IRange
{
    private readonly OoxmlSheet _sheet;

    internal OoxmlRange(OoxmlSheet sheet, int row1, int col1, int row2, int col2)
    {
        _sheet = sheet;
        FirstRow = Math.Min(row1, row2);
        LastRow = Math.Max(row1, row2);
        FirstCol = Math.Min(col1, col2);
        LastCol = Math.Max(col1, col2);
    }

    private OoxmlWorkbook Wb => _sheet.WorkbookInternal;

    public int FirstRow { get; }
    public int LastRow { get; }
    public int FirstCol { get; }
    public int LastCol { get; }

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

    private static void ApplyValue(ICell cell, object value)
    {
        switch (value)
        {
            case string s: cell.SetString(s); break;
            case bool b: cell.SetBool(b); break;
            case byte n: cell.SetNumber((double)n); break;
            case sbyte n: cell.SetNumber((double)n); break;
            case short n: cell.SetNumber((double)n); break;
            case int n: cell.SetNumber(n); break;
            case long n: cell.SetNumber(n); break;
            case float n: cell.SetNumber((double)n); break;
            case double n: cell.SetNumber(n); break;
            case decimal n: cell.SetNumber(n); break;
            case DateTime dt: cell.SetDate(dt); break;
            case DateOnly d: cell.SetDate(d); break;
            case TimeOnly t: cell.SetTime(t); break;
            case TimeSpan ts: cell.SetDuration(ts); break;
            default:
                throw new ArgumentException(
                    $"Unsupported value type '{value.GetType()}' for IRange.Value. " +
                    "Supported: string, bool, numeric (byte/sbyte/short/int/long/float/double/decimal), " +
                    "DateTime, DateOnly, TimeOnly, TimeSpan, or null to clear.",
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
