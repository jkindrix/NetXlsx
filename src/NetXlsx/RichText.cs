// Rich-text value types per design §6.8.1 / decision I-50.
// Excel's OOXML run model only supports per-run *font* properties
// (Bold, Italic, Underline, FontName, FontSize, Color) — not fills,
// borders, alignment, or number formats. RichTextStyle is therefore
// a font-only subset of CellStyle; cell-level visual style remains
// in CellStyle and is set via ICell.Style.

using System;
using System.Collections.Generic;
using System.Linq;

namespace NetXlsx;

/// <summary>
/// Font properties for a single rich-text run (decision I-50). A subset
/// of <see cref="CellStyle"/>: Excel's run model only honors font
/// properties, not fills/borders/alignment/number-format. Null axes
/// inherit from the cell's default font.
/// </summary>
public sealed record RichTextStyle
{
    /// <summary>The empty run style — every axis null. Inherits the cell's default font.</summary>
    public static RichTextStyle Default { get; } = new();

    /// <summary>Bold weight.</summary>
    public bool? Bold { get; init; }
    /// <summary>Italic.</summary>
    public bool? Italic { get; init; }
    /// <summary>Underline style.</summary>
    public UnderlineStyle? Underline { get; init; }
    /// <summary>Font family name (e.g. <c>"Calibri"</c>).</summary>
    public string? FontName { get; init; }
    /// <summary>Font size in points.</summary>
    public double? FontSize { get; init; }
    /// <summary>Font color. ARGB equality (decision I-23).</summary>
    public Color? Color { get; init; }

    /// <summary>
    /// Theme-based run color (decision I-89, mirroring
    /// <see cref="CellStyle.BackgroundTheme"/>). When set, takes
    /// precedence over <see cref="Color"/> and is written as the OOXML
    /// theme-index + tint color. The first theme-indexed write into a
    /// workbook without a theme part embeds
    /// <see cref="Workbook.DefaultThemeXml"/>.
    /// </summary>
    public ThemeColor? ColorTheme { get; init; }
}

/// <summary>
/// One contiguous run of formatted text inside a <see cref="RichText"/>.
/// <paramref name="Text"/> is the run's literal characters;
/// <paramref name="Style"/> describes how those characters are rendered.
/// An empty <paramref name="Text"/> is permitted but contributes no
/// formatting run to the cell.
/// </summary>
public sealed record RichTextRun(string Text, RichTextStyle Style)
{
    /// <summary>Constructs a run with the empty/default style.</summary>
    public RichTextRun(string text) : this(text, RichTextStyle.Default) { }
}

/// <summary>
/// A rich (multi-run) string value for a cell. Constructed once and
/// applied via <see cref="ICell.SetRichText"/>. Read back via
/// <see cref="ICell.GetRichText"/>; returns <c>null</c> for cells set
/// via plain <see cref="ICell.SetString"/>.
/// </summary>
public sealed record RichText
{
    /// <summary>The runs, in order. At least one run is required.</summary>
    public IReadOnlyList<RichTextRun> Runs { get; }

    /// <summary>Constructs a <see cref="RichText"/> from an explicit list of runs.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="runs"/> is null or contains a null run.</exception>
    /// <exception cref="ArgumentException"><paramref name="runs"/> is empty or any run has a null <c>Text</c>.</exception>
    public RichText(IReadOnlyList<RichTextRun> runs)
    {
        ArgumentNullException.ThrowIfNull(runs);
        if (runs.Count == 0)
            throw new ArgumentException("RichText must contain at least one run.", nameof(runs));
        for (int i = 0; i < runs.Count; i++)
        {
            var r = runs[i];
            if (r is null)
                throw new ArgumentNullException(nameof(runs), $"Run at index {i} is null.");
            if (r.Text is null)
                throw new ArgumentException($"Run at index {i} has null Text.", nameof(runs));
        }
        Runs = runs;
    }

    /// <summary>Constructs a <see cref="RichText"/> from a params array of runs.</summary>
    public RichText(params RichTextRun[] runs) : this((IReadOnlyList<RichTextRun>)runs) { }

    /// <summary>Concatenation of every run's <see cref="RichTextRun.Text"/>.</summary>
    public string PlainText => string.Concat(Runs.Select(r => r.Text));

    /// <summary>Structural equality: equal sequence of runs (per record equality on each).</summary>
    public bool Equals(RichText? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        if (Runs.Count != other.Runs.Count) return false;
        for (int i = 0; i < Runs.Count; i++)
            if (!Runs[i].Equals(other.Runs[i])) return false;
        return true;
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        var h = new HashCode();
        foreach (var r in Runs) h.Add(r);
        return h.ToHashCode();
    }
}
