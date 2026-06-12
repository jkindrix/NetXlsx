// I-82 engine swap — Open XML SDK-backed style pool (cell styles slice).
//
// This is the SDK counterpart to the NPOI-coupled CellStylePool. The dedup
// *logic* is engine-agnostic (decision #4 / #29): equal CellStyle values map to
// one cellXfs index, so a workbook that paints a million identically-styled
// cells allocates exactly one <xf>. Only the emission target changes — this pool
// writes OOXML schema types (CT_Font / CT_Fill / CT_Border / CT_Xf) directly
// instead of NPOI's XSSFCellStyle.
//
// Stylesheet structure maintained here (the order is the OOXML schema sequence):
//   numFmts? -> fonts -> fills -> borders -> cellStyleXfs -> cellXfs -> cellStyles
// On Create the pool builds a minimal valid stylesheet with the Excel
// conventions baked in (fills[0]=none, fills[1]=gray125, the Normal cellStyleXfs
// master per I-78, font[0] = the workbook's default font). On Open it adopts
// the file's existing stylesheet untouched and appends after it — the file's
// fonts/fills/styles and its default font (lesson #8) are preserved verbatim.
//
// Dedup honesty (the I-82 gotcha): the pools below track only the styles THIS
// pool allocates, exactly like CellStylePool. Pre-existing file <xf>s are left in
// place and never deduped against — a newly-applied style identical to a file
// style appends a fresh entry. That mirrors the NPOI engine's behavior and keeps
// existing cell -> style references valid.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlStylePool
{
    private readonly S.Stylesheet _ss;
    private readonly WorkbookOptions _options;

    // Dedup caches — NetXlsx-allocated entries only (see header note).
    private readonly Dictionary<CellStyle, uint> _xfPool = new();
    private readonly Dictionary<FontKey, uint> _fontPool = new();
    private readonly Dictionary<FillKey, uint> _fillPool = new();
    private readonly Dictionary<CellBorders, uint> _borderPool = new();
    private readonly Dictionary<string, uint> _numFmtPool = new(StringComparer.Ordinal);

    // Append cursors — initialized from the adopted/created stylesheet's counts.
    private uint _fontCount;
    private uint _fillCount;
    private uint _borderCount;
    private uint _cellXfCount;
    private uint _nextCustomNumFmtId = FirstCustomNumFmtId;

    // Diagnostics (decision I-61). Single-threaded under the workbook's own
    // mutation discipline, so plain ints suffice.
    internal int StyleHitCount;
    internal int StyleMissCount;
    internal int FontHitCount;
    internal int FontMissCount;
    internal int UniqueStyles => _xfPool.Count;
    internal int UniqueFonts => _fontPool.Count;

    // ---- I-89 lazy default-theme participation -------------------------------
    // Every theme-color ALLOCATION in this pool (a fill/font/border element
    // that carries a theme index) routes through NoteThemeColorWrite. The DOM
    // engine wires OnThemeColorWrite to OoxmlWorkbook.EnsureThemePart (the
    // single choke point); the streaming engine leaves the hook null and reads
    // UsesThemeColors at Save-time assembly instead (its stylesheet is detached
    // from any package while rows stream). Hits need no note: this pool dedups
    // only what it allocated, so every theme-carrying entry passed through the
    // miss path — and the hook — at least once.
    //
    // The one deliberate non-site: BuildDefaultStylesheet's font 0 carries
    // <color theme="1"/> as Excel-conventional scaffolding in EVERY created
    // workbook; treating it as a theme write would make the lazy embed eager
    // and break the byte-identity guarantee for theme-free workbooks.
    internal Action? OnThemeColorWrite { get; set; }
    internal bool UsesThemeColors { get; private set; }

    private void NoteThemeColorWrite()
    {
        UsesThemeColors = true;
        OnThemeColorWrite?.Invoke();
    }

    private const uint FirstCustomNumFmtId = 164;

    private OoxmlStylePool(S.Stylesheet ss, WorkbookOptions options)
    {
        _ss = ss;
        _options = options;
    }

    // ---- Construction -------------------------------------------------------

    /// <summary>
    /// Returns the workbook's style pool, materializing the WorkbookStylesPart and
    /// its stylesheet on first use. On a freshly created workbook this writes the
    /// Excel default scaffolding (and applies the workbook's default font to font
    /// index 0); on an opened workbook it adopts the file's stylesheet untouched.
    /// </summary>
    internal static OoxmlStylePool Attach(WorkbookPart wbPart, WorkbookOptions options)
    {
        var stylesPart = wbPart.WorkbookStylesPart;
        if (stylesPart is null)
        {
            stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = BuildDefaultStylesheet(options);
        }
        else if (stylesPart.Stylesheet is null || stylesPart.Stylesheet.GetFirstChild<S.CellFormats>() is null)
        {
            // A styles part with no usable cellXfs — treat as if absent.
            stylesPart.Stylesheet = BuildDefaultStylesheet(options);
        }

        var pool = new OoxmlStylePool(stylesPart.Stylesheet!, options);
        pool.InitializeCursors();
        return pool;
    }

    /// <summary>
    /// Builds a pool over a detached default stylesheet — no package exists yet.
    /// Used by the streaming engine (slice 9), which allocates cellXfs indices
    /// while rows are written forward-only and attaches the finished stylesheet
    /// to the WorkbookStylesPart only at Save-time assembly.
    /// </summary>
    internal static OoxmlStylePool CreateDetached(WorkbookOptions options)
    {
        var pool = new OoxmlStylePool(BuildDefaultStylesheet(options), options);
        pool.InitializeCursors();
        return pool;
    }

    /// <summary>The pool's stylesheet (for Save-time attachment of a detached pool).</summary>
    internal S.Stylesheet Stylesheet => _ss;

    private static S.Stylesheet BuildDefaultStylesheet(WorkbookOptions options)
    {
        var defaultFont = new S.Font(
            new S.FontSize { Val = options.DefaultFontSize },
            new S.Color { Theme = 1u },
            new S.FontName { Val = options.DefaultFontName },
            new S.FontFamilyNumbering { Val = 2 },
            new S.FontScheme { Val = S.FontSchemeValues.Minor });

        return new S.Stylesheet(
            new S.Fonts(defaultFont) { Count = 1u },
            new S.Fills(
                new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None }),
                new S.Fill(new S.PatternFill { PatternType = S.PatternValues.Gray125 })) { Count = 2u },
            new S.Borders(
                new S.Border(
                    new S.LeftBorder(), new S.RightBorder(), new S.TopBorder(),
                    new S.BottomBorder(), new S.DiagonalBorder())) { Count = 1u },
            // The Normal master xf (I-78). Excel resolves the default column width
            // from this style's font metrics; without it, column widths display wrong.
            new S.CellStyleFormats(
                new S.CellFormat { NumberFormatId = 0u, FontId = 0u, FillId = 0u, BorderId = 0u }) { Count = 1u },
            new S.CellFormats(
                new S.CellFormat { NumberFormatId = 0u, FontId = 0u, FillId = 0u, BorderId = 0u, FormatId = 0u }) { Count = 1u },
            new S.CellStyles(
                new S.CellStyle { Name = "Normal", FormatId = 0u, BuiltinId = 0u }) { Count = 1u });
    }

    private void InitializeCursors()
    {
        _fontCount = (uint)(_ss.GetFirstChild<S.Fonts>()?.Elements<S.Font>().Count() ?? 0);
        _fillCount = (uint)(_ss.GetFirstChild<S.Fills>()?.Elements<S.Fill>().Count() ?? 0);
        _borderCount = (uint)(_ss.GetFirstChild<S.Borders>()?.Elements<S.Border>().Count() ?? 0);
        _cellXfCount = (uint)(_ss.GetFirstChild<S.CellFormats>()?.Elements<S.CellFormat>().Count() ?? 0);

        // Continue custom numFmt ids after the highest one already present so we
        // never collide with the file's existing custom formats.
        var numFmts = _ss.GetFirstChild<S.NumberingFormats>();
        if (numFmts is not null)
        {
            foreach (var nf in numFmts.Elements<S.NumberingFormat>())
            {
                if (nf.NumberFormatId?.Value is uint id && id >= _nextCustomNumFmtId)
                    _nextCustomNumFmtId = id + 1;
                if (nf.FormatCode?.Value is { } code && nf.NumberFormatId?.Value is uint existing)
                    _numFmtPool[code] = existing;
            }
        }
    }

    // ---- Allocation ---------------------------------------------------------

    // Apply-memo (I-87): bulk styled writes re-apply the same immutable
    // CellStyle instances over and over (the documented palette pattern), and
    // each apply pays a Merge record allocation plus a record value-hash on
    // the dedup lookup. Since CellStyle is a sealed record with init-only
    // state, (same instance, same starting xf index) always merges to the
    // same result — memoize the apply by reference. The cap bounds growth for
    // the anti-pattern of allocating a fresh CellStyle per cell (those
    // entries can never hit). Cleared with the row caches on every
    // escape-hatch access (an out-of-band stylesheet edit is the same
    // coherence class as an out-of-band grid edit).
    private const int ApplyMemoCap = 1024;
    private readonly Dictionary<CellStyle, (uint From, uint To)> _applyMemo =
        new(ReferenceEqualityComparer.Instance);

    internal bool TryMemoizedApply(CellStyle style, uint from, out uint to)
    {
        if (_applyMemo.TryGetValue(style, out var m) && m.From == from)
        {
            // A memo hit IS a pool hit — the application reused an existing
            // pooled entry; only the merge/dedup work was skipped. Keeps the
            // public StyleHitCount diagnostic contract intact.
            StyleHitCount++;
            to = m.To;
            return true;
        }
        to = 0u;
        return false;
    }

    internal void MemoizeApply(CellStyle style, uint from, uint to)
    {
        if (_applyMemo.Count >= ApplyMemoCap) _applyMemo.Clear();
        _applyMemo[style] = (from, to);
    }

    internal void ClearApplyMemo() => _applyMemo.Clear();

    /// <summary>
    /// Resolves <paramref name="style"/> to a cellXfs index, allocating a new
    /// <c>&lt;xf&gt;</c> (and its font/fill/border/numFmt) only when no
    /// structurally-equal style has been pooled. The all-null style maps to
    /// index 0 (the workbook default) without allocating.
    /// </summary>
    internal uint GetOrCreate(CellStyle style)
    {
        if (IsEmpty(style)) return 0u;

        if (_xfPool.TryGetValue(style, out var existing))
        {
            StyleHitCount++;
            return existing;
        }
        StyleMissCount++;

        uint fontId = NeedsFont(style) ? GetOrCreateFont(FontKey.From(style)) : 0u;
        uint fillId = GetOrCreateFill(style);
        uint borderId = style.Borders is { } b ? GetOrCreateBorder(b) : 0u;
        uint numFmtId = !string.IsNullOrEmpty(style.NumberFormat) ? GetOrCreateNumFmt(style.NumberFormat!) : 0u;

        var xf = new S.CellFormat
        {
            NumberFormatId = numFmtId,
            FontId = fontId,
            FillId = fillId,
            BorderId = borderId,
            FormatId = 0u,
            ApplyNumberFormat = numFmtId != 0 ? true : null,
            ApplyFont = fontId != 0 ? true : null,
            ApplyFill = fillId != 0 ? true : null,
            ApplyBorder = borderId != 0 ? true : null,
        };

        var alignment = BuildAlignment(style);
        if (alignment is not null)
        {
            xf.ApplyAlignment = true;
            xf.Alignment = alignment;
        }

        var cellXfs = _ss.GetFirstChild<S.CellFormats>()!;
        cellXfs.AppendChild(xf);
        uint index = _cellXfCount++;
        cellXfs.Count = _cellXfCount;

        _xfPool[style] = index;
        return index;
    }

    private static S.Alignment? BuildAlignment(CellStyle style)
    {
        if (style.HorizontalAlignment is null && style.VerticalAlignment is null && style.WrapText is null)
            return null;
        var a = new S.Alignment();
        if (style.HorizontalAlignment is { } h) a.Horizontal = MapHAlign(h);
        if (style.VerticalAlignment is { } v) a.Vertical = MapVAlign(v);
        if (style.WrapText == true) a.WrapText = true;
        return a;
    }

    private uint GetOrCreateFont(FontKey key)
    {
        if (_fontPool.TryGetValue(key, out var existing))
        {
            FontHitCount++;
            return existing;
        }
        FontMissCount++;

        // Theme font color takes precedence over explicit RGB (the I-79 rule,
        // verbatim from the fill path) and participates in the I-89 lazy embed.
        if (key.ColorTheme is not null) NoteThemeColorWrite();

        var font = new S.Font();
        if (key.Bold) font.AppendChild(new S.Bold());
        if (key.Italic) font.AppendChild(new S.Italic());
        if (key.Underline != UnderlineStyle.None) font.AppendChild(new S.Underline { Val = MapUnderline(key.Underline) });
        font.AppendChild(new S.FontSize { Val = key.Size ?? _options.DefaultFontSize });
        if (key.ColorTheme is { } theme) font.AppendChild(ThemeToColor<S.Color>(theme));
        else if (key.Color is { } c) font.AppendChild(ToColor<S.Color>(c));
        font.AppendChild(new S.FontName { Val = key.Name ?? _options.DefaultFontName });

        var fonts = _ss.GetFirstChild<S.Fonts>()!;
        fonts.AppendChild(font);
        uint index = _fontCount++;
        fonts.Count = _fontCount;

        _fontPool[key] = index;
        return index;
    }

    private uint GetOrCreateFill(CellStyle style)
    {
        // Theme background takes precedence over explicit RGB (decision I-79):
        // theme + tint preserves Excel's exact tint rendering.
        FillKey key;
        if (style.BackgroundTheme is { } theme) key = FillKey.OfTheme(theme);
        else if (style.Background is { } rgb) key = FillKey.OfRgb(rgb);
        else return 0u; // no fill -> default (none)

        if (_fillPool.TryGetValue(key, out var existing)) return existing;

        if (key.Theme is not null) NoteThemeColorWrite(); // I-89 lazy embed

        var fg = new S.ForegroundColor();
        if (key.Theme is { } t)
        {
            fg.Theme = (uint)t.Index;
            if (t.Tint != 0) fg.Tint = t.Tint;
        }
        else
        {
            fg.Rgb = ArgbHex(key.Rgb!.Value);
        }

        var fill = new S.Fill(new S.PatternFill(fg, new S.BackgroundColor { Indexed = 64u })
        {
            PatternType = S.PatternValues.Solid,
        });

        var fills = _ss.GetFirstChild<S.Fills>()!;
        fills.AppendChild(fill);
        uint index = _fillCount++;
        fills.Count = _fillCount;

        _fillPool[key] = index;
        return index;
    }

    private uint GetOrCreateBorder(CellBorders b)
    {
        if (_borderPool.TryGetValue(b, out var existing)) return existing;

        // I-89 lazy embed — fire once when any edge carries a theme color.
        // Theme colors on style-less edges are inert (BuildEdge writes a color
        // only when the edge has a BorderStyle, the literal-color contract),
        // so only edges that will actually emit count.
        if ((b.Left is not null and not BorderStyle.None && b.LeftColorTheme is not null)
            || (b.Right is not null and not BorderStyle.None && b.RightColorTheme is not null)
            || (b.Top is not null and not BorderStyle.None && b.TopColorTheme is not null)
            || (b.Bottom is not null and not BorderStyle.None && b.BottomColorTheme is not null))
        {
            NoteThemeColorWrite();
        }

        var border = new S.Border(
            BuildEdge<S.LeftBorder>(b.Left, b.LeftColor, b.LeftColorTheme),
            BuildEdge<S.RightBorder>(b.Right, b.RightColor, b.RightColorTheme),
            BuildEdge<S.TopBorder>(b.Top, b.TopColor, b.TopColorTheme),
            BuildEdge<S.BottomBorder>(b.Bottom, b.BottomColor, b.BottomColorTheme),
            new S.DiagonalBorder());

        var borders = _ss.GetFirstChild<S.Borders>()!;
        borders.AppendChild(border);
        uint index = _borderCount++;
        borders.Count = _borderCount;

        _borderPool[b] = index;
        return index;
    }

    private static TEdge BuildEdge<TEdge>(BorderStyle? style, Color? color, ThemeColor? colorTheme)
        where TEdge : S.BorderPropertiesType, new()
    {
        var edge = new TEdge();
        if (style is { } s && s != BorderStyle.None)
        {
            edge.Style = MapBorder(s);
            // Theme wins over the literal color when both are set (the I-79
            // precedence rule, per edge — decision I-89).
            if (colorTheme is { } t) edge.AppendChild(ThemeToColor<S.Color>(t));
            else if (color is { } c) edge.AppendChild(ToColor<S.Color>(c));
        }
        return edge;
    }

    private uint GetOrCreateNumFmt(string formatCode)
    {
        // Reuse a builtin id where one matches, so e.g. "0.00" round-trips as the
        // builtin rather than a redundant custom entry.
        if (BuiltinNumFmtId(formatCode) is { } builtin) return builtin;

        if (_numFmtPool.TryGetValue(formatCode, out var existing)) return existing;

        var numFmts = _ss.GetFirstChild<S.NumberingFormats>();
        if (numFmts is null)
        {
            numFmts = new S.NumberingFormats { Count = 0u };
            _ss.InsertAt(numFmts, 0); // numFmts is the first stylesheet child
        }

        uint id = _nextCustomNumFmtId++;
        numFmts.AppendChild(new S.NumberingFormat { NumberFormatId = id, FormatCode = formatCode });
        numFmts.Count = (uint)numFmts.Elements<S.NumberingFormat>().Count();

        _numFmtPool[formatCode] = id;
        return id;
    }

    // ---- Read-back (xf index -> CellStyle) ----------------------------------

    /// <summary>
    /// Parses the cellXfs entry at <paramref name="xfIndex"/> back to a
    /// <see cref="CellStyle"/> value. Index 0 (and any out-of-range index) is the
    /// workbook default and reads as <see cref="CellStyle.Default"/>. Lossy for
    /// axes NetXlsx does not model — same contract as the NPOI engine's reader.
    /// </summary>
    internal CellStyle ReadStyle(uint xfIndex)
    {
        if (xfIndex == 0) return CellStyle.Default;
        var cellXfs = _ss.GetFirstChild<S.CellFormats>();
        var xf = cellXfs?.Elements<S.CellFormat>().ElementAtOrDefault((int)xfIndex);
        if (xf is null) return CellStyle.Default;
        return ReadXf(xf);
    }

    // Parses one CT_Xf back to a CellStyle. Shared by the cellXfs read path
    // (ReadStyle) and the cellStyleXfs read path (ReadNamedStyles) — the two
    // tables carry the same element type.
    private CellStyle ReadXf(S.CellFormat xf)
    {
        var fontStyle = ReadFont(xf.FontId?.Value ?? 0u);
        return new CellStyle
        {
            Bold = fontStyle.Bold,
            Italic = fontStyle.Italic,
            Underline = fontStyle.Underline,
            FontName = fontStyle.FontName,
            FontSize = fontStyle.FontSize,
            FontColor = fontStyle.FontColor,
            FontColorTheme = fontStyle.FontColorTheme,
            Background = ReadFillRgb(xf.FillId?.Value ?? 0u),
            BackgroundTheme = ReadFillTheme(xf.FillId?.Value ?? 0u),
            NumberFormat = ReadNumFmt(xf.NumberFormatId?.Value ?? 0u),
            HorizontalAlignment = xf.Alignment?.Horizontal is { } h ? MapHAlignBack(h.Value) : null,
            VerticalAlignment = xf.Alignment?.Vertical is { } v ? MapVAlignBack(v.Value) : null,
            WrapText = xf.Alignment?.WrapText?.Value == true ? true : null,
            Borders = ReadBorder(xf.BorderId?.Value ?? 0u),
        };
    }

    /// <summary>Returns the number-format code for the cellXfs entry, or null.</summary>
    internal string? NumberFormatOf(uint xfIndex)
    {
        var cellXfs = _ss.GetFirstChild<S.CellFormats>();
        var xf = cellXfs?.Elements<S.CellFormat>().ElementAtOrDefault((int)xfIndex);
        return xf is null ? null : ReadNumFmt(xf.NumberFormatId?.Value ?? 0u);
    }

    /// <summary>
    /// The effective font + layout axes IColumn.AutoSize measures with for a
    /// cellXfs entry (I-84): font axes resolved through the xf's font element
    /// with font-0 fallback, plus alignment indent and the raw OOXML text
    /// rotation (0–180; NPOI feeds the raw value into its trig, so we do too).
    /// </summary>
    internal AutoSizeStyle AutoSizeStyleOf(uint xfIndex)
    {
        var xf = _ss.GetFirstChild<S.CellFormats>()?.Elements<S.CellFormat>().ElementAtOrDefault((int)xfIndex);
        var (name, size, bold, italic) = EffectiveFontAxes(xf?.FontId?.Value ?? 0u);
        int indent = (int)(xf?.Alignment?.Indent?.Value ?? 0u);
        int rotation = (int)(xf?.Alignment?.TextRotation?.Value ?? 0u);
        return new AutoSizeStyle(name, size, bold, italic, indent, rotation);
    }

    // A font element missing a name/size axis inherits font 0's (the workbook
    // default); the literal Calibri/11 floor only matters for malformed files
    // with no usable font 0.
    private (string Name, double Size, bool Bold, bool Italic) EffectiveFontAxes(uint fontId)
    {
        var fonts = _ss.GetFirstChild<S.Fonts>();
        var font = fonts?.Elements<S.Font>().ElementAtOrDefault((int)fontId);
        var font0 = fontId == 0 ? font : fonts?.Elements<S.Font>().ElementAtOrDefault(0);
        string name = font?.GetFirstChild<S.FontName>()?.Val?.Value
            ?? font0?.GetFirstChild<S.FontName>()?.Val?.Value
            ?? "Calibri";
        double size = font?.GetFirstChild<S.FontSize>()?.Val?.Value
            ?? font0?.GetFirstChild<S.FontSize>()?.Val?.Value
            ?? 11.0;
        bool bold = font?.GetFirstChild<S.Bold>() is { } b && (b.Val?.Value ?? true);
        bool italic = font?.GetFirstChild<S.Italic>() is { } i && (i.Val?.Value ?? true);
        return (name, size, bold, italic);
    }

    private (bool? Bold, bool? Italic, UnderlineStyle? Underline, string? FontName, double? FontSize, Color? FontColor, ThemeColor? FontColorTheme)
        ReadFont(uint fontId)
    {
        var font = _ss.GetFirstChild<S.Fonts>()?.Elements<S.Font>().ElementAtOrDefault((int)fontId);
        if (font is null) return (null, null, null, null, null, null, null);

        bool? bold = font.GetFirstChild<S.Bold>() is { } bx ? (bx.Val?.Value ?? true) ? true : (bool?)null : null;
        bool? italic = font.GetFirstChild<S.Italic>() is { } ix ? (ix.Val?.Value ?? true) ? true : (bool?)null : null;
        UnderlineStyle? underline = font.GetFirstChild<S.Underline>() is { } ux
            ? MapUnderlineBack(ux.Val?.Value ?? S.UnderlineValues.Single)
            : null;
        string? name = font.GetFirstChild<S.FontName>()?.Val?.Value;
        double? size = font.GetFirstChild<S.FontSize>()?.Val?.Value;
        // A theme-indexed color reads via FontColorTheme with the literal axis
        // null, and vice versa (decision I-89 — same exclusivity as the fill
        // path's ReadFillRgb/ReadFillTheme pair).
        var colorEl = font.GetFirstChild<S.Color>();
        ThemeColor? colorTheme = FromColorTheme(colorEl);
        Color? color = colorTheme is null ? FromColor(colorEl) : null;
        return (bold, italic, underline, name, size, color, colorTheme);
    }

    private S.PatternFill? PatternFillAt(uint fillId)
        => _ss.GetFirstChild<S.Fills>()?.Elements<S.Fill>().ElementAtOrDefault((int)fillId)?.PatternFill;

    private Color? ReadFillRgb(uint fillId)
    {
        var pf = PatternFillAt(fillId);
        if (pf?.PatternType?.Value != S.PatternValues.Solid) return null;
        var fg = pf.ForegroundColor;
        // A theme-backed fill is reported via BackgroundTheme, not RGB.
        if (fg?.Theme is not null) return null;
        return FromColor(fg);
    }

    private ThemeColor? ReadFillTheme(uint fillId)
    {
        var pf = PatternFillAt(fillId);
        if (pf?.PatternType?.Value != S.PatternValues.Solid) return null;
        var fg = pf.ForegroundColor;
        if (fg?.Theme?.Value is not { } theme) return null;
        return new ThemeColor((int)theme, fg.Tint?.Value ?? 0);
    }

    private CellBorders? ReadBorder(uint borderId)
    {
        if (borderId == 0) return null;
        var border = _ss.GetFirstChild<S.Borders>()?.Elements<S.Border>().ElementAtOrDefault((int)borderId);
        if (border is null) return null;

        var top = MapBorderBack(border.TopBorder);
        var right = MapBorderBack(border.RightBorder);
        var bottom = MapBorderBack(border.BottomBorder);
        var left = MapBorderBack(border.LeftBorder);
        if (top is null && right is null && bottom is null && left is null) return null;

        // Per edge: a theme-indexed color reads via the theme property with
        // the literal axis null, and vice versa (decision I-89).
        var topTheme = FromColorTheme(border.TopBorder?.Color);
        var rightTheme = FromColorTheme(border.RightBorder?.Color);
        var bottomTheme = FromColorTheme(border.BottomBorder?.Color);
        var leftTheme = FromColorTheme(border.LeftBorder?.Color);
        return new CellBorders(
            top, topTheme is null ? FromColor(border.TopBorder?.Color) : null,
            right, rightTheme is null ? FromColor(border.RightBorder?.Color) : null,
            bottom, bottomTheme is null ? FromColor(border.BottomBorder?.Color) : null,
            left, leftTheme is null ? FromColor(border.LeftBorder?.Color) : null)
        {
            TopColorTheme = topTheme,
            RightColorTheme = rightTheme,
            BottomColorTheme = bottomTheme,
            LeftColorTheme = leftTheme,
        };
    }

    private string? ReadNumFmt(uint numFmtId)
    {
        if (numFmtId == 0) return null;
        if (BuiltinFormatCode(numFmtId) is { } builtin) return builtin;
        var nf = _ss.GetFirstChild<S.NumberingFormats>()?
            .Elements<S.NumberingFormat>()
            .FirstOrDefault(x => x.NumberFormatId?.Value == numFmtId);
        return nf?.FormatCode?.Value;
    }

    // ---- Date-format detection (for ICell.GetDate / Kind) -------------------

    /// <summary>
    /// True when the cellXfs entry at <paramref name="xfIndex"/> carries a
    /// date/time number format — the OOXML equivalent of NPOI's
    /// <c>DateUtil.IsCellDateFormatted</c>. Builtin ids 14–22 / 45–47 are dates,
    /// and any custom format code with date tokens (y/m/d/h/s) is a date.
    /// </summary>
    internal bool IsDateFormatted(uint xfIndex)
    {
        var cellXfs = _ss.GetFirstChild<S.CellFormats>();
        var xf = cellXfs?.Elements<S.CellFormat>().ElementAtOrDefault((int)xfIndex);
        if (xf is null) return false;
        uint numFmtId = xf.NumberFormatId?.Value ?? 0u;
        if (IsBuiltinDateId(numFmtId)) return true;
        if (numFmtId < FirstCustomNumFmtId && numFmtId != 0) return false; // non-date builtin
        var code = ReadNumFmt(numFmtId);
        return code is not null && IsDateFormatCode(code);
    }

    private static bool IsBuiltinDateId(uint id)
        => (id >= 14 && id <= 22) || id == 45 || id == 46 || id == 47;

    // Mirrors NPOI's IsADateFormat string heuristic: strip color/condition
    // brackets, locale prefixes, escaped chars and quoted literals, then look for
    // a date/time token. Elapsed-time brackets [h]/[m]/[s] survive and count.
    internal static bool IsDateFormatCode(string format)
    {
        if (string.IsNullOrWhiteSpace(format)) return false;
        if (format.Equals("general", StringComparison.OrdinalIgnoreCase)) return false;

        var sb = new System.Text.StringBuilder(format.Length);
        for (int i = 0; i < format.Length; i++)
        {
            char ch = format[i];
            switch (ch)
            {
                case '\\': i++; break;                 // escaped next char — skip both
                case '"':                               // quoted literal — skip to close
                    i++;
                    while (i < format.Length && format[i] != '"') i++;
                    break;
                case '[':                               // [Red]/[$-409]/[>0] vs elapsed [h]
                    int close = format.IndexOf(']', i);
                    if (close < 0) { i = format.Length; break; }
                    string inner = format.Substring(i + 1, close - i - 1);
                    char c0 = inner.Length > 0 ? char.ToLowerInvariant(inner[0]) : '\0';
                    if (c0 == 'h' || c0 == 'm' || c0 == 's') sb.Append(c0); // elapsed token
                    i = close;
                    break;
                default:
                    sb.Append(ch);
                    break;
            }
        }

        foreach (char ch in sb.ToString())
        {
            switch (char.ToLowerInvariant(ch))
            {
                case 'y': case 'd': case 'h': case 's': return true;
                case 'm': return true; // month or minute — either is a date/time token
            }
        }
        return false;
    }

    // ---- Rich-text run properties (<rPr>) -----------------------------------
    // A rich-text run carries its font as inline <rPr> properties on the <r>
    // node, NOT a cellXfs font index — so these build/read S.RunProperties
    // directly and never touch the font pool. The marquee semantic (lesson #10):
    // a run whose style is empty gets NO <rPr>, so it inherits the cell's font.
    // Absent axes inside a present <rPr> inherit too — so, unlike the cell-style
    // font path, a null FontName/FontSize is omitted rather than defaulted.

    /// <summary>
    /// Builds the <c>&lt;rPr&gt;</c> for a rich-text run, or <c>null</c> when the
    /// run style is empty (the run then inherits the cell font — lesson #10). The
    /// child order mirrors the styles <c>&lt;font&gt;</c> sequence
    /// (b/i/u/sz/color/rFont) Excel writes; the run-font element is
    /// <c>&lt;rFont&gt;</c>, not <c>&lt;name&gt;</c>.
    /// </summary>
    internal static S.RunProperties? BuildRunProperties(RichTextStyle style)
    {
        if (IsEmptyRunStyle(style)) return null;

        var rpr = new S.RunProperties();
        if (style.Bold == true) rpr.AppendChild(new S.Bold());
        if (style.Italic == true) rpr.AppendChild(new S.Italic());
        if (style.Underline is { } u && u != UnderlineStyle.None)
            rpr.AppendChild(new S.Underline { Val = MapUnderline(u) });
        if (style.FontSize is { } sz) rpr.AppendChild(new S.FontSize { Val = sz });
        // Theme wins over the literal color when both are set (the I-79
        // precedence rule — decision I-89). The I-89 lazy embed for this path
        // is OoxmlCell.SetRichText's EnsureThemePart call: this builder is
        // static (no pool instance) and the run lives in cell XML, not the
        // stylesheet, so the pool hook cannot cover it.
        if (style.ColorTheme is { } ct) rpr.AppendChild(ThemeToColor<S.Color>(ct));
        else if (style.Color is { } c) rpr.AppendChild(ToColor<S.Color>(c));
        if (!string.IsNullOrEmpty(style.FontName)) rpr.AppendChild(new S.RunFont { Val = style.FontName });
        return rpr;
    }

    /// <summary>
    /// Reads a rich-text run's <c>&lt;rPr&gt;</c> back to a
    /// <see cref="RichTextStyle"/>. A <c>null</c> <paramref name="rpr"/> (no run
    /// properties) reads as <see cref="RichTextStyle.Default"/> — the inherited
    /// cell font. Lossy for axes the run model does not cover, exactly like
    /// <see cref="ReadFont"/>; a theme-indexed run color reads via
    /// <see cref="RichTextStyle.ColorTheme"/> with the literal axis null, and
    /// vice versa (decision I-89; indexed-palette colors still read null).
    /// </summary>
    internal static RichTextStyle ReadRunStyle(S.RunProperties? rpr)
    {
        if (rpr is null) return RichTextStyle.Default;

        bool? bold = rpr.GetFirstChild<S.Bold>() is { } bx ? (bx.Val?.Value ?? true) ? true : (bool?)null : null;
        bool? italic = rpr.GetFirstChild<S.Italic>() is { } ix ? (ix.Val?.Value ?? true) ? true : (bool?)null : null;
        UnderlineStyle? underline = rpr.GetFirstChild<S.Underline>() is { } ux
            ? MapUnderlineBack(ux.Val?.Value ?? S.UnderlineValues.Single)
            : null;
        string? name = rpr.GetFirstChild<S.RunFont>()?.Val?.Value;
        double? size = rpr.GetFirstChild<S.FontSize>()?.Val?.Value;
        var colorEl = rpr.GetFirstChild<S.Color>();
        ThemeColor? colorTheme = FromColorTheme(colorEl);
        Color? color = colorTheme is null ? FromColor(colorEl) : null;

        return new RichTextStyle
        {
            Bold = bold,
            Italic = italic,
            Underline = underline,
            FontName = name,
            FontSize = size,
            Color = color,
            ColorTheme = colorTheme,
        };
    }

    private static bool IsEmptyRunStyle(RichTextStyle s) =>
        s.Bold is null && s.Italic is null && s.Underline is null
        && s.FontName is null && s.FontSize is null && s.Color is null
        && s.ColorTheme is null;

    // ---- Differential formats (<dxfs>, for conditional formatting) ----------
    //
    // A cfRule's style is a dxf (differential format) in styles.xml, referenced
    // by @dxfId. Only the axes the NPOI engine's ApplyStyle honors are modeled
    // (decision I-73): Bold / Italic (a <font> with flag elements) and
    // Background (a solid <patternFill> with fgColor + bgColor indexed 64 —
    // the NPOI shape, proven to render in Excel). Unlike NPOI, which allocates
    // a fresh dxf per rule (the oracle dump shows three identical fills),
    // structurally equal dxfs dedup to one entry per the pool's #4 discipline.

    private readonly Dictionary<(bool Bold, bool Italic, Color? Background), uint> _dxfPool = new();

    internal uint GetOrCreateDifferentialFormat(CellStyle style)
    {
        var key = (Bold: style.Bold == true, Italic: style.Italic == true, style.Background);
        if (_dxfPool.TryGetValue(key, out uint existing)) return existing;

        var dxf = new S.DifferentialFormat();
        if (key.Bold || key.Italic)
        {
            var font = new S.Font();
            if (key.Bold) font.AppendChild(new S.Bold());
            if (key.Italic) font.AppendChild(new S.Italic());
            dxf.AppendChild(font);
        }
        if (style.Background is { } bg)
        {
            dxf.AppendChild(new S.Fill(new S.PatternFill(
                ToColor<S.ForegroundColor>(bg),
                new S.BackgroundColor { Indexed = 64u })
            { PatternType = S.PatternValues.Solid }));
        }

        var dxfs = GetOrCreateDxfs();
        dxfs.AppendChild(dxf);
        uint id = (uint)dxfs.Elements<S.DifferentialFormat>().Count() - 1;
        dxfs.Count = id + 1;
        _dxfPool[key] = id;
        return id;
    }

    private S.DifferentialFormats GetOrCreateDxfs()
    {
        var existing = _ss.GetFirstChild<S.DifferentialFormats>();
        if (existing is not null) return existing;
        // CT_Stylesheet order: … cellStyles, dxfs, tableStyles, colors, extLst —
        // insert before the first trailing sibling an opened file may carry
        // (quirk #3: the stylesheet is order-constrained and AppendChild does
        // not reorder).
        var dxfs = new S.DifferentialFormats { Count = 0u };
        var successor = _ss.ChildElements.FirstOrDefault(e =>
            e is S.TableStyles || e is S.Colors || e is S.StylesheetExtensionList);
        if (successor is null) _ss.AppendChild(dxfs);
        else _ss.InsertBefore(dxfs, successor);
        return dxfs;
    }

    // ---- Named styles (the NPOI engine's I-67 round-trip) --------------------
    //
    // RegisterStyle persists each name as a cellStyleXfs <xf> (mirroring the
    // deduped cellXfs entry's component ids) plus a <cellStyle name=… xfId=…>
    // entry, so registered names survive a save/open round-trip and appear in
    // Excel's Cell Styles panel. Divergence from the NPOI witness (oracle-dumped
    // 2026-06-03): NPOI stamps builtinId="0" on every user entry — but builtinId
    // 0 *claims the Normal builtin* per ECMA-376; user styles carry no builtinId
    // here (Excel's own files don't either). Re-registering a name repoints its
    // xfId and leaves the superseded cellStyleXfs entry orphaned — harmless, and
    // exactly what NPOI does.

    /// <summary>Upserts the named-style OOXML entry for <paramref name="name"/>.</summary>
    internal void WriteNamedStyle(string name, CellStyle style)
    {
        // The style's component ids come from the same dedup allocation the
        // visual style uses (an empty style maps to the Normal components).
        uint xfIndex = GetOrCreate(style);
        var sourceXf = _ss.GetFirstChild<S.CellFormats>()!
            .Elements<S.CellFormat>().ElementAt((int)xfIndex);

        var namedXf = new S.CellFormat
        {
            NumberFormatId = sourceXf.NumberFormatId?.Value ?? 0u,
            FontId = sourceXf.FontId?.Value ?? 0u,
            FillId = sourceXf.FillId?.Value ?? 0u,
            BorderId = sourceXf.BorderId?.Value ?? 0u,
            ApplyNumberFormat = sourceXf.ApplyNumberFormat?.Value == true ? true : null,
            ApplyFont = sourceXf.ApplyFont?.Value == true ? true : null,
            ApplyFill = sourceXf.ApplyFill?.Value == true ? true : null,
            ApplyBorder = sourceXf.ApplyBorder?.Value == true ? true : null,
        };
        if (sourceXf.Alignment is not null)
        {
            namedXf.ApplyAlignment = true;
            namedXf.Alignment = (S.Alignment)sourceXf.Alignment.CloneNode(true);
        }

        var styleXfs = GetOrCreateCellStyleFormats();
        styleXfs.AppendChild(namedXf);
        uint xfId = (uint)styleXfs.Elements<S.CellFormat>().Count() - 1;
        styleXfs.Count = xfId + 1;

        var cellStyles = GetOrCreateCellStyles();
        var existing = cellStyles.Elements<S.CellStyle>().FirstOrDefault(cs =>
            string.Equals(cs.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            existing.FormatId = xfId;
        }
        else
        {
            cellStyles.AppendChild(new S.CellStyle { Name = name, FormatId = xfId });
        }
        cellStyles.Count = (uint)cellStyles.Elements<S.CellStyle>().Count();
    }

    /// <summary>
    /// Reads the persisted named-style entries back as (name, style) pairs —
    /// the rehydration source for a freshly opened workbook's registry. The
    /// built-in "Normal" entry and nameless entries are skipped (they are not
    /// user-registered names; same contract as the NPOI engine).
    /// </summary>
    internal IEnumerable<KeyValuePair<string, CellStyle>> ReadNamedStyles()
    {
        var cellStyles = _ss.GetFirstChild<S.CellStyles>();
        if (cellStyles is null) yield break;
        var styleXfs = _ss.GetFirstChild<S.CellStyleFormats>();

        foreach (var cs in cellStyles.Elements<S.CellStyle>())
        {
            var name = cs.Name?.Value;
            if (string.IsNullOrEmpty(name)) continue;
            if (string.Equals(name, "Normal", StringComparison.OrdinalIgnoreCase)) continue;

            var xf = styleXfs?.Elements<S.CellFormat>()
                .ElementAtOrDefault((int)(cs.FormatId?.Value ?? 0u));
            if (xf is null) continue;

            yield return new KeyValuePair<string, CellStyle>(name, ReadXf(xf));
        }
    }

    private S.CellStyleFormats GetOrCreateCellStyleFormats()
    {
        var existing = _ss.GetFirstChild<S.CellStyleFormats>();
        if (existing is not null) return existing;
        // CT_Stylesheet order: … borders, cellStyleXfs, cellXfs, … — created
        // stylesheets always carry the Normal master xf (I-78), but an opened
        // file's stylesheet may omit the table entirely. Seed the fresh table
        // with the Normal master at index 0: cellXfs entries reference their
        // parent via xfId, and the universal convention (this engine's created
        // files included) is xfId=0 = Normal — a named xf landing at index 0
        // would silently re-parent every styled cell.
        var styleXfs = new S.CellStyleFormats(
            new S.CellFormat { NumberFormatId = 0u, FontId = 0u, FillId = 0u, BorderId = 0u }) { Count = 1u };
        var successor = _ss.ChildElements.FirstOrDefault(e =>
            e is S.CellFormats || e is S.CellStyles || e is S.DifferentialFormats
            || e is S.TableStyles || e is S.Colors || e is S.StylesheetExtensionList);
        if (successor is null) _ss.AppendChild(styleXfs);
        else _ss.InsertBefore(styleXfs, successor);
        return styleXfs;
    }

    private S.CellStyles GetOrCreateCellStyles()
    {
        var existing = _ss.GetFirstChild<S.CellStyles>();
        if (existing is not null) return existing;
        // CT_Stylesheet order: … cellXfs, cellStyles, dxfs, … (quirk #3). A
        // fresh table is seeded with the builtin Normal entry (xfId 0) — the
        // panel Excel expects, and CT_CellStyles requires ≥1 child (quirk #7).
        var cellStyles = new S.CellStyles(
            new S.CellStyle { Name = "Normal", FormatId = 0u, BuiltinId = 0u }) { Count = 1u };
        var successor = _ss.ChildElements.FirstOrDefault(e =>
            e is S.DifferentialFormats || e is S.TableStyles || e is S.Colors
            || e is S.StylesheetExtensionList);
        if (successor is null) _ss.AppendChild(cellStyles);
        else _ss.InsertBefore(cellStyles, successor);
        return cellStyles;
    }

    // ---- Helpers ------------------------------------------------------------

    private static bool IsEmpty(CellStyle s) =>
        s.Bold is null && s.Italic is null && s.Underline is null
        && s.FontName is null && s.FontSize is null && s.FontColor is null
        && s.FontColorTheme is null
        && s.Background is null && s.BackgroundTheme is null
        && s.NumberFormat is null
        && s.HorizontalAlignment is null && s.VerticalAlignment is null
        && s.WrapText is null && s.Borders is null;

    private static bool NeedsFont(CellStyle s) =>
        s.Bold is not null || s.Italic is not null || s.Underline is not null
        || s.FontName is not null || s.FontSize is not null || s.FontColor is not null
        || s.FontColorTheme is not null;

    private static TColor ToColor<TColor>(Color c) where TColor : S.ColorType, new()
        => new() { Rgb = ArgbHex(c) };

    // Theme-indexed CT_Color emission (decision I-89): theme index + tint,
    // tint omitted when zero — the same convention the fill path established
    // with I-79.
    private static TColor ThemeToColor<TColor>(ThemeColor t) where TColor : S.ColorType, new()
    {
        var color = new TColor { Theme = (uint)t.Index };
        if (t.Tint != 0) color.Tint = t.Tint;
        return color;
    }

    private static Color? FromColor(S.ColorType? c)
    {
        if (c?.Rgb?.Value is not { } hex) return null;
        // ARGB hex ("FFRRGGBB") or RGB hex ("RRGGBB").
        if (hex.Length == 8) hex = hex.Substring(2);
        if (hex.Length != 6) return null;
        if (!byte.TryParse(hex.AsSpan(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte r)
            || !byte.TryParse(hex.AsSpan(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte g)
            || !byte.TryParse(hex.AsSpan(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte b))
            return null;
        return Color.FromRgb(r, g, b);
    }

    // Theme-indexed CT_Color read-back (decision I-89): theme index + tint,
    // matching ReadFillTheme's shape. Null for literal/indexed colors.
    private static ThemeColor? FromColorTheme(S.ColorType? c)
        => c?.Theme?.Value is { } theme ? new ThemeColor((int)theme, c.Tint?.Value ?? 0) : null;

    private static string ArgbHex(Color c)
        => $"FF{c.R:X2}{c.G:X2}{c.B:X2}";

    // ---- Enum maps (mirror CellStylePool's NPOI maps) -----------------------

    private static S.HorizontalAlignmentValues MapHAlign(HAlign h) => h switch
    {
        HAlign.General => S.HorizontalAlignmentValues.General,
        HAlign.Left => S.HorizontalAlignmentValues.Left,
        HAlign.Center => S.HorizontalAlignmentValues.Center,
        HAlign.Right => S.HorizontalAlignmentValues.Right,
        HAlign.Fill => S.HorizontalAlignmentValues.Fill,
        HAlign.Justify => S.HorizontalAlignmentValues.Justify,
        _ => S.HorizontalAlignmentValues.General,
    };

    private static HAlign? MapHAlignBack(S.HorizontalAlignmentValues h)
    {
        if (h == S.HorizontalAlignmentValues.Left) return HAlign.Left;
        if (h == S.HorizontalAlignmentValues.Center) return HAlign.Center;
        if (h == S.HorizontalAlignmentValues.Right) return HAlign.Right;
        if (h == S.HorizontalAlignmentValues.Fill) return HAlign.Fill;
        if (h == S.HorizontalAlignmentValues.Justify) return HAlign.Justify;
        return null; // General == "inherit/default"
    }

    private static S.VerticalAlignmentValues MapVAlign(VAlign v) => v switch
    {
        VAlign.Top => S.VerticalAlignmentValues.Top,
        VAlign.Center => S.VerticalAlignmentValues.Center,
        VAlign.Bottom => S.VerticalAlignmentValues.Bottom,
        VAlign.Justify => S.VerticalAlignmentValues.Justify,
        _ => S.VerticalAlignmentValues.Bottom,
    };

    private static VAlign? MapVAlignBack(S.VerticalAlignmentValues v)
    {
        if (v == S.VerticalAlignmentValues.Top) return VAlign.Top;
        if (v == S.VerticalAlignmentValues.Center) return VAlign.Center;
        if (v == S.VerticalAlignmentValues.Justify) return VAlign.Justify;
        return null; // Bottom == Excel default == "no explicit value"
    }

    private static S.BorderStyleValues MapBorder(BorderStyle b) => b switch
    {
        BorderStyle.None => S.BorderStyleValues.None,
        BorderStyle.Thin => S.BorderStyleValues.Thin,
        BorderStyle.Medium => S.BorderStyleValues.Medium,
        BorderStyle.Thick => S.BorderStyleValues.Thick,
        BorderStyle.Double => S.BorderStyleValues.Double,
        BorderStyle.Dashed => S.BorderStyleValues.Dashed,
        BorderStyle.Dotted => S.BorderStyleValues.Dotted,
        _ => S.BorderStyleValues.None,
    };

    private static BorderStyle? MapBorderBack(S.BorderPropertiesType? edge)
    {
        if (edge?.Style is not { } s) return null;
        var v = s.Value;
        if (v == S.BorderStyleValues.None) return null;
        if (v == S.BorderStyleValues.Thin) return BorderStyle.Thin;
        if (v == S.BorderStyleValues.Medium) return BorderStyle.Medium;
        if (v == S.BorderStyleValues.Thick) return BorderStyle.Thick;
        if (v == S.BorderStyleValues.Double) return BorderStyle.Double;
        if (v == S.BorderStyleValues.Dashed) return BorderStyle.Dashed;
        if (v == S.BorderStyleValues.Dotted) return BorderStyle.Dotted;
        return BorderStyle.Thin; // unmapped styles fall through as Thin (parity with NPOI reader)
    }

    private static S.UnderlineValues MapUnderline(UnderlineStyle u) => u switch
    {
        UnderlineStyle.None => S.UnderlineValues.None,
        UnderlineStyle.Single => S.UnderlineValues.Single,
        UnderlineStyle.Double => S.UnderlineValues.Double,
        UnderlineStyle.SingleAccounting => S.UnderlineValues.SingleAccounting,
        UnderlineStyle.DoubleAccounting => S.UnderlineValues.DoubleAccounting,
        _ => S.UnderlineValues.None,
    };

    private static UnderlineStyle? MapUnderlineBack(S.UnderlineValues u)
    {
        if (u == S.UnderlineValues.Single) return UnderlineStyle.Single;
        if (u == S.UnderlineValues.Double) return UnderlineStyle.Double;
        if (u == S.UnderlineValues.SingleAccounting) return UnderlineStyle.SingleAccounting;
        if (u == S.UnderlineValues.DoubleAccounting) return UnderlineStyle.DoubleAccounting;
        return null; // None
    }

    // ---- Builtin number formats ---------------------------------------------
    // The subset relevant to NetXlsx's surface — enough for round-trip of the
    // formats we emit and detection of common builtin date ids on opened files.

    private static uint? BuiltinNumFmtId(string formatCode) => formatCode switch
    {
        "General" => 0u,
        "0" => 1u,
        "0.00" => 2u,
        "#,##0" => 3u,
        "#,##0.00" => 4u,
        "0%" => 9u,
        "0.00%" => 10u,
        "m/d/yy" => 14u,
        "d-mmm-yy" => 15u,
        "d-mmm" => 16u,
        "mmm-yy" => 17u,
        "h:mm AM/PM" => 18u,
        "h:mm:ss AM/PM" => 19u,
        "h:mm" => 20u,
        "h:mm:ss" => 21u,
        "m/d/yy h:mm" => 22u,
        "mm:ss" => 45u,
        "[h]:mm:ss" => 46u,
        "mm:ss.0" => 47u,
        "@" => 49u,
        _ => null,
    };

    private static string? BuiltinFormatCode(uint id) => id switch
    {
        1u => "0",
        2u => "0.00",
        3u => "#,##0",
        4u => "#,##0.00",
        9u => "0%",
        10u => "0.00%",
        14u => "m/d/yy",
        15u => "d-mmm-yy",
        16u => "d-mmm",
        17u => "mmm-yy",
        18u => "h:mm AM/PM",
        19u => "h:mm:ss AM/PM",
        20u => "h:mm",
        21u => "h:mm:ss",
        22u => "m/d/yy h:mm",
        45u => "mm:ss",
        46u => "[h]:mm:ss",
        47u => "mm:ss.0",
        49u => "@",
        _ => null,
    };

    // ---- Dedup keys ---------------------------------------------------------

    private readonly record struct FontKey(string? Name, double? Size, bool Bold, bool Italic, UnderlineStyle Underline, Color? Color, ThemeColor? ColorTheme)
    {
        internal static FontKey From(CellStyle s) => new(
            s.FontName, s.FontSize, s.Bold ?? false, s.Italic ?? false,
            s.Underline ?? UnderlineStyle.None, s.FontColor, s.FontColorTheme);
    }

    private readonly record struct FillKey(Color? Rgb, ThemeColor? Theme)
    {
        internal static FillKey OfRgb(Color c) => new(c, null);
        internal static FillKey OfTheme(ThemeColor t) => new(null, t);
    }
}
