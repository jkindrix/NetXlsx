// Sheet-protection options record per design §6.4.3 / decision I-53.
// Composed with ISheet.Protect to apply granular permission flags.

namespace NetXlsx;

/// <summary>
/// Granular sheet-protection options for <see cref="ISheet.Protect"/>.
/// Each <c>Lock*</c> property, when <c>true</c>, prevents the
/// corresponding user action even while the sheet is protected.
/// Defaults (<see cref="Default"/>) leave every action permitted —
/// the user can still see protected cells; cells marked locked
/// (the Excel default) still cannot be edited.
/// <para>
/// Excel sheet protection is a <b>UX guard</b>, not real security:
/// the password (when supplied) is hashed with a weak algorithm
/// long known to be brute-forceable. Use for "stop accidental
/// edits", not "stop a determined attacker."
/// </para>
/// </summary>
public sealed record SheetProtection
{
    /// <summary>The empty protection — every action permitted (most permissive).</summary>
    public static SheetProtection Default { get; } = new();

    /// <summary>A maximally-restrictive protection — every <c>Lock*</c> flag set.</summary>
    public static SheetProtection LockAll { get; } = new()
    {
        LockFormatCells = true,
        LockFormatColumns = true,
        LockFormatRows = true,
        LockInsertColumns = true,
        LockInsertRows = true,
        LockInsertHyperlinks = true,
        LockDeleteColumns = true,
        LockDeleteRows = true,
        LockSelectLockedCells = true,
        LockSelectUnlockedCells = true,
        LockSort = true,
        LockAutoFilter = true,
        LockPivotTables = true,
        LockObjects = true,
        LockScenarios = true,
    };

    /// <summary>Lock cell formatting.</summary>
    public bool LockFormatCells { get; init; }
    /// <summary>Lock column formatting (width, hide/unhide).</summary>
    public bool LockFormatColumns { get; init; }
    /// <summary>Lock row formatting (height, hide/unhide).</summary>
    public bool LockFormatRows { get; init; }
    /// <summary>Lock column insertion.</summary>
    public bool LockInsertColumns { get; init; }
    /// <summary>Lock row insertion.</summary>
    public bool LockInsertRows { get; init; }
    /// <summary>Lock hyperlink insertion.</summary>
    public bool LockInsertHyperlinks { get; init; }
    /// <summary>Lock column deletion.</summary>
    public bool LockDeleteColumns { get; init; }
    /// <summary>Lock row deletion.</summary>
    public bool LockDeleteRows { get; init; }
    /// <summary>Lock selection of locked cells (hides them from selection).</summary>
    public bool LockSelectLockedCells { get; init; }
    /// <summary>Lock selection of unlocked cells.</summary>
    public bool LockSelectUnlockedCells { get; init; }
    /// <summary>Lock sort.</summary>
    public bool LockSort { get; init; }
    /// <summary>Lock AutoFilter use.</summary>
    public bool LockAutoFilter { get; init; }
    /// <summary>Lock pivot-table manipulation.</summary>
    public bool LockPivotTables { get; init; }
    /// <summary>Lock embedded-object editing (charts, pictures, shapes).</summary>
    public bool LockObjects { get; init; }
    /// <summary>Lock scenario manipulation.</summary>
    public bool LockScenarios { get; init; }
}
