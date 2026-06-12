// Cell style value record + supporting enums per design §6.7 / §6.8.
// CellStyle equality is structural; the style-pool dedup (decision #4)
// keys stylesheet <xf> (cellXfs) entries on CellStyle value equality.

namespace NetXlsx;

/// <summary>
/// Immutable description of a cell's visual style. Equality is structural:
/// two <see cref="CellStyle"/> values with the same property values share
/// one underlying stylesheet <c>&lt;xf&gt;</c> entry via the workbook's
/// style-pool dedup (decision #4). Properties typed as <c>Nullable{T}</c>: <c>null</c>
/// means "inherit existing" when applied via <see cref="ICell.Style"/>;
/// a non-null value overwrites the cell's current style on that axis.
/// </summary>
public sealed record CellStyle
{
    /// <summary>The empty style — all properties null. Applied to a cell, leaves the cell unchanged.</summary>
    public static CellStyle Default { get; } = new();

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
    public Color? FontColor { get; init; }

    /// <summary>
    /// Theme-based font color (decision I-89, mirroring
    /// <see cref="BackgroundTheme"/>). When set, takes precedence over
    /// <see cref="FontColor"/> and is written as the OOXML theme-index +
    /// tint color, preserving Excel's exact rendering when a theme is
    /// present. The first theme-indexed write into a workbook without a
    /// theme part embeds <see cref="Workbook.DefaultThemeXml"/>.
    /// </summary>
    public ThemeColor? FontColorTheme { get; init; }

    /// <summary>Solid-fill background color. ARGB equality.</summary>
    public Color? Background { get; init; }

    /// <summary>
    /// Theme-based background color (decision I-79). When set, takes
    /// precedence over <see cref="Background"/> and is written as the
    /// OOXML theme-index + tint color. This preserves Excel's exact
    /// color rendering when a theme is present, since explicit RGB
    /// doesn't always match Excel's tint calculation. The first
    /// theme-indexed write into a workbook without a theme part embeds
    /// <see cref="Workbook.DefaultThemeXml"/> (decision I-89).
    /// </summary>
    public ThemeColor? BackgroundTheme { get; init; }

    /// <summary>Excel number format string (e.g. <c>"$#,##0.00"</c>, <c>"yyyy-mm-dd"</c>). Pass-through bytes per §7.2.</summary>
    public string? NumberFormat { get; init; }

    /// <summary>Horizontal alignment.</summary>
    public HAlign? HorizontalAlignment { get; init; }
    /// <summary>Vertical alignment.</summary>
    public VAlign? VerticalAlignment { get; init; }

    /// <summary>Whether the cell wraps long text.</summary>
    public bool? WrapText { get; init; }

    /// <summary>Border styles.</summary>
    public CellBorders? Borders { get; init; }
}

/// <summary>
/// Cell border styles per edge. The theme-color properties (decision I-89,
/// mirroring <see cref="CellStyle.BackgroundTheme"/>) are init-only — set
/// them via an object initializer or <c>with</c> expression. Per edge, the
/// theme variant takes precedence over the literal color when both are
/// set; either is written only when that edge also has a
/// <see cref="BorderStyle"/>. The first theme-indexed write into a
/// workbook without a theme part embeds
/// <see cref="Workbook.DefaultThemeXml"/>.
/// </summary>
public sealed record CellBorders(
    BorderStyle? Top = null,    Color? TopColor = null,
    BorderStyle? Right = null,  Color? RightColor = null,
    BorderStyle? Bottom = null, Color? BottomColor = null,
    BorderStyle? Left = null,   Color? LeftColor = null)
{
    /// <summary>Theme-based color for the top edge; wins over <see cref="TopColor"/>.</summary>
    public ThemeColor? TopColorTheme { get; init; }
    /// <summary>Theme-based color for the right edge; wins over <see cref="RightColor"/>.</summary>
    public ThemeColor? RightColorTheme { get; init; }
    /// <summary>Theme-based color for the bottom edge; wins over <see cref="BottomColor"/>.</summary>
    public ThemeColor? BottomColorTheme { get; init; }
    /// <summary>Theme-based color for the left edge; wins over <see cref="LeftColor"/>.</summary>
    public ThemeColor? LeftColorTheme { get; init; }

    /// <summary>Build a uniform border on all four edges.</summary>
    public static CellBorders All(BorderStyle style, Color? color = null) =>
        new(style, color, style, color, style, color, style, color);
}

/// <summary>Border line styles (mirrors Excel's set).</summary>
public enum BorderStyle
{
    /// <summary>No border.</summary>
    None,
    /// <summary>Thin solid line.</summary>
    Thin,
    /// <summary>Medium-weight solid line.</summary>
    Medium,
    /// <summary>Thick solid line.</summary>
    Thick,
    /// <summary>Double line.</summary>
    Double,
    /// <summary>Dashed line.</summary>
    Dashed,
    /// <summary>Dotted line.</summary>
    Dotted,
}

/// <summary>Horizontal alignment.</summary>
public enum HAlign
{
    /// <summary>General (Excel default — left for text, right for numbers).</summary>
    General,
    /// <summary>Left.</summary>
    Left,
    /// <summary>Center.</summary>
    Center,
    /// <summary>Right.</summary>
    Right,
    /// <summary>Fill the cell with the value repeated.</summary>
    Fill,
    /// <summary>Justify multi-line content.</summary>
    Justify,
}

/// <summary>Vertical alignment.</summary>
public enum VAlign
{
    /// <summary>Top.</summary>
    Top,
    /// <summary>Center.</summary>
    Center,
    /// <summary>Bottom (Excel default).</summary>
    Bottom,
    /// <summary>Justify multi-line content vertically.</summary>
    Justify,
}

/// <summary>Underline styles.</summary>
public enum UnderlineStyle
{
    /// <summary>No underline.</summary>
    None,
    /// <summary>Single underline.</summary>
    Single,
    /// <summary>Double underline.</summary>
    Double,
    /// <summary>Single accounting underline.</summary>
    SingleAccounting,
    /// <summary>Double accounting underline.</summary>
    DoubleAccounting,
}

/// <summary>
/// A theme-based color reference (decision I-79). The <see cref="Index"/>
/// references one of the workbook's theme colors (0–11: dark1, light1,
/// dark2, light2, accent1–6, hyperlink, followedHyperlink). The optional
/// <see cref="Tint"/> applies a lightening (positive) or darkening
/// (negative) modifier in [-1.0, 1.0].
/// <para>
/// Use this when reproducing files that use Excel's theme colors —
/// explicit RGB doesn't always match Excel's tint calculation, so theme
/// references preserve exact rendering when the workbook has a theme.
/// A workbook that has no theme part receives the standard Office theme
/// (<see cref="Workbook.DefaultThemeXml"/>) automatically on the first
/// theme-indexed styling write (decision I-89), so theme references
/// resolve consistently across consumers.
/// </para>
/// </summary>
public sealed record ThemeColor(int Index, double Tint = 0);
