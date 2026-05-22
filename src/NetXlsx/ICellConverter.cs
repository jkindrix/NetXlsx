// Custom type converters per design §6.9.1 / decision I-58.
// Plugged into the typed-row mapping via ColumnAttribute.ConverterType.

namespace NetXlsx;

/// <summary>
/// Translates a single cell between its Excel representation and a
/// user-defined type <typeparamref name="T"/> in the typed-row mapping
/// (<c>[Worksheet]</c> source-generator pipeline).
/// <para>
/// Implementations are referenced from <see cref="ColumnAttribute.ConverterType"/>;
/// the source generator instantiates one per property (cached as a
/// <c>static readonly</c> field on the generated extension class) and
/// dispatches read/write through it.
/// </para>
/// <para>
/// Implementations must have a public parameterless constructor. The
/// generator allocates the instance once at class-init time, so
/// per-call allocations should be avoided inside
/// <see cref="Write"/> / <see cref="Read"/>.
/// </para>
/// </summary>
public interface ICellConverter<T>
{
    /// <summary>
    /// Writes <paramref name="value"/> to <paramref name="cell"/>. The
    /// implementation is responsible for picking the correct
    /// <c>ICell.Set*</c> overload (or composing multiple cells via
    /// <c>ICell.Underlying</c> if the OOXML representation is exotic —
    /// rare).
    /// </summary>
    void Write(ICell cell, T value);

    /// <summary>
    /// Reads the value stored in <paramref name="cell"/> back to a
    /// <typeparamref name="T"/>. Implementations should throw a
    /// descriptive exception (<see cref="WorkbookException"/> or
    /// derived) when the cell's kind is incompatible.
    /// </summary>
    T Read(ICell cell);
}
