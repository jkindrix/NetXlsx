// IRange — rectangular range surface (decision §6.10). Sparse vs dense
// enumeration, Apply/Value/ClearContents, Merge. Split out of ISheet.cs
// at v1.2 / v1.1-review item 2.

using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace NetXlsx;

/// <summary>
/// A rectangular range of cells on an <see cref="ISheet"/>.
/// Enumeration is sparse by default (only currently-populated cells);
/// use <see cref="EnumerateAll"/> for dense iteration.
/// </summary>
public interface IRange : System.Collections.Generic.IEnumerable<ICell>
{
    /// <summary>
    /// Canonical A1 form of the range — <c>A1:C3</c> for bounded
    /// ranges, single-cell form for 1×1 ranges. Whole-row and
    /// whole-column shorthand expands to the canonical bounded form
    /// per design §6.10.
    /// </summary>
    string Address { get; }

    /// <summary>1-based top row.</summary>
    int FirstRow { get; }
    /// <summary>1-based bottom row (inclusive).</summary>
    int LastRow { get; }
    /// <summary>1-based leftmost column.</summary>
    int FirstCol { get; }
    /// <summary>1-based rightmost column (inclusive).</summary>
    int LastCol { get; }

    /// <summary>The total number of cell coordinates in the rectangle (dense count).</summary>
    int Count { get; }

    /// <summary>The owning sheet.</summary>
    ISheet Sheet { get; }

    /// <summary>
    /// Yields every cell coordinate in the rectangle, including blanks.
    /// Lazily materializes empty cells on demand. For whole-row /
    /// whole-column ranges this can be 1M+ items — sparse base
    /// enumeration via <c>foreach</c> is the usual idiom.
    /// </summary>
    System.Collections.Generic.IEnumerable<ICell> EnumerateAll();

    /// <summary>
    /// Sets every cell in the rectangle to <paramref name="value"/>.
    /// Dispatched on runtime type: string / bool / numeric (int, long,
    /// float, double, decimal) / DateTime / DateOnly / TimeOnly /
    /// TimeSpan. Null clears the cell. Unsupported types throw
    /// <see cref="System.ArgumentException"/>.
    /// <para>
    /// <b>Performance / precision note (decision #5):</b> this overload
    /// is a convenience wrapper. It boxes value types and dispatches at
    /// runtime via <c>is</c> checks. For tight loops, large workloads
    /// (&gt;10k cells), or cases where the value type is known
    /// statically, prefer the per-type setters
    /// (<see cref="ICell.SetString"/> / <see cref="ICell.SetNumber(double)"/>
    /// / etc.) on each cell — they avoid the boxing and dispatch cost,
    /// and for <c>decimal</c> they document the IEEE-754 precision
    /// trade-off explicitly (§7.4). The convenience overload is
    /// intentional and supported; the typed setters are intentional
    /// and faster.
    /// </para>
    /// </summary>
    IRange Value(object? value);

    /// <summary>Applies <paramref name="style"/> to every cell in the rectangle (dense).</summary>
    IRange Apply(CellStyle style);

    /// <summary>
    /// Applies the style registered under <paramref name="name"/> to every
    /// cell in the rectangle (decision I-57). See
    /// <see cref="IWorkbook.RegisterStyle"/> for the name registry.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="name"/> is null.</exception>
    /// <exception cref="ArgumentException">No style is registered under <paramref name="name"/>.</exception>
    IRange ApplyNamedStyle(string name);

    /// <summary>
    /// Merges this range. Shorthand for
    /// <c>sheet.MergeCells(range.Address)</c>; same semantics
    /// (decision §6.4): 1×1 is a no-op; overlap with an existing merge
    /// throws.
    /// </summary>
    IRange Merge();

    /// <summary>
    /// Clears the value of every cell in the rectangle. Styles are
    /// preserved. Inverse of <see cref="Value"/> with a non-null arg.
    /// </summary>
    IRange ClearContents();
}
