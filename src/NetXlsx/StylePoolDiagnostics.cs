// Style-pool diagnostics per design §6.2.4 / decision I-61.
// Read-only counters over the workbook's CellStyle + Font dedup
// pools. Useful for ops verification ("is the dedup actually
// working in production?") and as a sanity check in tests.

namespace NetXlsx;

/// <summary>
/// Snapshot of <see cref="IWorkbook"/>'s style-pool dedup activity
/// (decision I-61). Returned by
/// <see cref="IWorkbook.GetStylePoolDiagnostics"/> as a value-type
/// snapshot — calling it does not allocate, and the counters do
/// not update after the snapshot is taken.
/// <para>
/// The pools sit behind decision #4: equal <see cref="CellStyle"/>
/// values share one underlying NPOI <c>ICellStyle</c>, and fonts
/// with identical properties share one <c>IFont</c>. Hit / miss
/// counters track how often this dedup kicked in vs allocated a
/// new entry.
/// </para>
/// </summary>
public readonly struct StylePoolDiagnostics
{
    /// <summary>Number of <see cref="CellStyle"/> lookups that returned an existing pool entry.</summary>
    public int StyleHitCount { get; }

    /// <summary>Number of <see cref="CellStyle"/> lookups that allocated a new pool entry.</summary>
    public int StyleMissCount { get; }

    /// <summary>Number of font lookups that returned an existing pool entry.</summary>
    public int FontHitCount { get; }

    /// <summary>Number of font lookups that allocated a new pool entry.</summary>
    public int FontMissCount { get; }

    /// <summary>Current size of the <see cref="CellStyle"/> pool — distinct allocated styles.</summary>
    public int UniqueStyles { get; }

    /// <summary>Current size of the font pool — distinct allocated fonts.</summary>
    public int UniqueFonts { get; }

    /// <summary>
    /// Fraction of <see cref="CellStyle"/> lookups served from the pool —
    /// <c>hits / (hits + misses)</c>. Returns 0 when no lookups have
    /// occurred. A healthy workbook with reused styles trends toward 1.
    /// </summary>
    public double StyleDedupRatio =>
        (StyleHitCount + StyleMissCount) == 0
            ? 0
            : (double)StyleHitCount / (StyleHitCount + StyleMissCount);

    /// <summary>Fraction of font lookups served from the font pool.</summary>
    public double FontDedupRatio =>
        (FontHitCount + FontMissCount) == 0
            ? 0
            : (double)FontHitCount / (FontHitCount + FontMissCount);

    /// <summary>Constructs a diagnostics snapshot.</summary>
    public StylePoolDiagnostics(
        int styleHits, int styleMisses, int fontHits, int fontMisses,
        int uniqueStyles, int uniqueFonts)
    {
        StyleHitCount = styleHits;
        StyleMissCount = styleMisses;
        FontHitCount = fontHits;
        FontMissCount = fontMisses;
        UniqueStyles = uniqueStyles;
        UniqueFonts = uniqueFonts;
    }
}
