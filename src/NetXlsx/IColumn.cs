// IColumn — sheet-column handle (width, hidden, AutoSize, default style).
// Split out of ISheet.cs at v1.2 / v1.1-review item 2.

using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace NetXlsx;

/// <summary>
/// A single column on an <see cref="ISheet"/>. Columns are lightweight
/// handles — accessing a column does not materialize cells or mutate the
/// underlying file.
/// </summary>
public interface IColumn
{
    /// <summary>The 1-based column index (<c>"A" == 1</c>).</summary>
    int Index { get; }

    /// <summary>The canonical column letter (<c>1 → "A"</c>, <c>27 → "AA"</c>).</summary>
    string Letter { get; }

    /// <summary>The owning sheet.</summary>
    ISheet Sheet { get; }

    /// <summary>
    /// Whether this column is hidden in Excel. Setter takes effect
    /// immediately; reading reflects the current NPOI state.
    /// </summary>
    bool Hidden { get; set; }

    /// <summary>
    /// Column width in Excel "character" units. Setting writes through
    /// to NPOI's 256ths-of-a-character integer representation.
    /// </summary>
    double WidthUnits { get; set; }

    /// <summary>Fluent form of the <see cref="WidthUnits"/> setter.</summary>
    /// <exception cref="System.ArgumentOutOfRangeException"><paramref name="units"/> is negative or NaN.</exception>
    IColumn Width(double units);

    /// <summary>
    /// Sizes this column to fit its populated contents, measuring with
    /// embedded font-metric tables — deterministic on every machine,
    /// headless included (design decision I-84, which superseded the
    /// environment-dependent OS-font-stack behavior of decision I3).
    /// Throws <see cref="MissingFontException"/> only for fonts outside
    /// the embedded metric set, naming the font.
    /// </summary>
    /// <exception cref="MissingFontException">The column's fonts are not in the embedded metric set.</exception>
    IColumn AutoSize();

    /// <summary>
    /// Applies <paramref name="apply"/> to every populated cell in this
    /// column, top to bottom. Empty cells are skipped (sparse iteration).
    /// </summary>
    IColumn ForEachPopulated(Action<ICell> apply);

    /// <summary>
    /// Sets <paramref name="style"/> as the column's default style. New
    /// cells in this column inherit it; existing cells are unaffected
    /// (NPOI's <c>SetDefaultColumnStyle</c> semantics).
    /// </summary>
    IColumn SetDefaultStyle(CellStyle style);
}
