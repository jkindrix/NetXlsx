// IRow — 1-based row indexer + fluent per-column setters (decision §6.5).
// Split out of ISheet.cs at v1.2 / v1.1-review item 2.

using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace NetXlsx;

/// <summary>
/// Represents a row within an <see cref="ISheet"/>. Cells are 1-based
/// (decision #3); fluent setters return the row itself for chaining.
/// </summary>
public interface IRow
{
    /// <summary>The row's 1-based index.</summary>
    int Index { get; }

    /// <summary>The owning sheet.</summary>
    ISheet Sheet { get; }

    /// <summary>
    /// Returns the cell at <paramref name="column"/> (1-based), materializing
    /// an empty cell if none exists.
    /// </summary>
    /// <exception cref="System.ArgumentOutOfRangeException">Column out of range.</exception>
    ICell Cell(int column);

    /// <summary>Indexer form of <see cref="Cell(int)"/>.</summary>
    ICell this[int column] { get; }

    /// <summary>Indexer keyed by column letter (e.g. <c>"A"</c>, <c>"AA"</c>).</summary>
    /// <exception cref="System.ArgumentException">Not a valid column letter.</exception>
    ICell this[string columnLetter] { get; }

    /// <summary>Writes a string value to the column and returns this row for chaining.</summary>
    IRow Set(int column, string value);
    /// <summary>Writes a double value to the column and returns this row for chaining.</summary>
    IRow Set(int column, double value);
    /// <summary>Writes a decimal value to the column and returns this row for chaining.</summary>
    IRow Set(int column, decimal value);
    /// <summary>Writes an int value to the column and returns this row for chaining.</summary>
    IRow Set(int column, int value);
    /// <summary>Writes a long value to the column and returns this row for chaining.</summary>
    IRow Set(int column, long value);
    /// <summary>Writes a bool value to the column and returns this row for chaining.</summary>
    IRow Set(int column, bool value);
    /// <summary>Writes a <see cref="DateTime"/> value (decisions I-17, I-18).</summary>
    IRow Set(int column, DateTime value);
    /// <summary>Writes a <see cref="DateOnly"/> value (decision I-19).</summary>
    IRow Set(int column, DateOnly value);
    /// <summary>Writes a <see cref="TimeOnly"/> value as a fraction of a day.</summary>
    IRow Set(int column, TimeOnly value);
    /// <summary>Writes a <see cref="TimeSpan"/> value as elapsed time.</summary>
    /// <exception cref="ArgumentOutOfRangeException">Negative <paramref name="value"/> (decision I15).</exception>
    IRow Set(int column, TimeSpan value);

    /// <summary>Whether this row is hidden in Excel.</summary>
    bool Hidden { get; set; }

    /// <summary>
    /// Escape hatch — direct access to the underlying NPOI <c>XSSFRow</c>.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFRow Underlying { get; }
}
