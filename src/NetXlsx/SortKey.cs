using System;

namespace NetXlsx;

/// <summary>
/// Describes a single sort key for <see cref="ISheet.SortRange"/>
/// (decision I-72). Each key identifies a column (1-based) and a
/// direction. Multiple keys are applied in order (primary, secondary, …).
/// </summary>
public sealed class SortKey
{
    /// <summary>Column index (1-based) to sort by.</summary>
    public int Column { get; }

    /// <summary>Sort direction — <c>true</c> for ascending (A→Z, 0→9).</summary>
    public bool Ascending { get; }

    private SortKey(int column, bool ascending)
    {
        if (column < 1)
            throw new ArgumentOutOfRangeException(nameof(column), column, "Column must be >= 1.");
        Column = column;
        Ascending = ascending;
    }

    /// <summary>Sort ascending on the given column (1-based).</summary>
    public static SortKey Asc(int column) => new(column, true);

    /// <summary>Sort descending on the given column (1-based).</summary>
    public static SortKey Desc(int column) => new(column, false);
}
