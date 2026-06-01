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
// On CreateOoxml the pool builds a minimal valid stylesheet with the Excel
// conventions baked in (fills[0]=none, fills[1]=gray125, the Normal cellStyleXfs
// master per I-78, font[0] = the workbook's default font). On OpenOoxml it adopts
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

        var font = new S.Font();
        if (key.Bold) font.AppendChild(new S.Bold());
        if (key.Italic) font.AppendChild(new S.Italic());
        if (key.Underline != UnderlineStyle.None) font.AppendChild(new S.Underline { Val = MapUnderline(key.Underline) });
        font.AppendChild(new S.FontSize { Val = key.Size ?? _options.DefaultFontSize });
        if (key.Color is { } c) font.AppendChild(ToColor<S.Color>(c));
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

        var border = new S.Border(
            BuildEdge<S.LeftBorder>(b.Left, b.LeftColor),
            BuildEdge<S.RightBorder>(b.Right, b.RightColor),
            BuildEdge<S.TopBorder>(b.Top, b.TopColor),
            BuildEdge<S.BottomBorder>(b.Bottom, b.BottomColor),
            new S.DiagonalBorder());

        var borders = _ss.GetFirstChild<S.Borders>()!;
        borders.AppendChild(border);
        uint index = _borderCount++;
        borders.Count = _borderCount;

        _borderPool[b] = index;
        return index;
    }

    private static TEdge BuildEdge<TEdge>(BorderStyle? style, Color? color)
        where TEdge : S.BorderPropertiesType, new()
    {
        var edge = new TEdge();
        if (style is { } s && s != BorderStyle.None)
        {
            edge.Style = MapBorder(s);
            if (color is { } c) edge.AppendChild(ToColor<S.Color>(c));
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

        var fontStyle = ReadFont(xf.FontId?.Value ?? 0u);
        return new CellStyle
        {
            Bold = fontStyle.Bold,
            Italic = fontStyle.Italic,
            Underline = fontStyle.Underline,
            FontName = fontStyle.FontName,
            FontSize = fontStyle.FontSize,
            FontColor = fontStyle.FontColor,
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

    private (bool? Bold, bool? Italic, UnderlineStyle? Underline, string? FontName, double? FontSize, Color? FontColor)
        ReadFont(uint fontId)
    {
        var font = _ss.GetFirstChild<S.Fonts>()?.Elements<S.Font>().ElementAtOrDefault((int)fontId);
        if (font is null) return (null, null, null, null, null, null);

        bool? bold = font.GetFirstChild<S.Bold>() is { } bx ? (bx.Val?.Value ?? true) ? true : (bool?)null : null;
        bool? italic = font.GetFirstChild<S.Italic>() is { } ix ? (ix.Val?.Value ?? true) ? true : (bool?)null : null;
        UnderlineStyle? underline = font.GetFirstChild<S.Underline>() is { } ux
            ? MapUnderlineBack(ux.Val?.Value ?? S.UnderlineValues.Single)
            : null;
        string? name = font.GetFirstChild<S.FontName>()?.Val?.Value;
        double? size = font.GetFirstChild<S.FontSize>()?.Val?.Value;
        Color? color = FromColor(font.GetFirstChild<S.Color>());
        return (bold, italic, underline, name, size, color);
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

        return new CellBorders(
            top, FromColor(border.TopBorder?.Color),
            right, FromColor(border.RightBorder?.Color),
            bottom, FromColor(border.BottomBorder?.Color),
            left, FromColor(border.LeftBorder?.Color));
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
        if (style.Color is { } c) rpr.AppendChild(ToColor<S.Color>(c));
        if (!string.IsNullOrEmpty(style.FontName)) rpr.AppendChild(new S.RunFont { Val = style.FontName });
        return rpr;
    }

    /// <summary>
    /// Reads a rich-text run's <c>&lt;rPr&gt;</c> back to a
    /// <see cref="RichTextStyle"/>. A <c>null</c> <paramref name="rpr"/> (no run
    /// properties) reads as <see cref="RichTextStyle.Default"/> — the inherited
    /// cell font. Lossy for axes the run model does not cover, exactly like
    /// <see cref="ReadFont"/>; theme/indexed run colors read as <c>null</c> until
    /// the theme slice, matching the NPOI engine's RGB-only run-color read.
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
        Color? color = FromColor(rpr.GetFirstChild<S.Color>());

        return new RichTextStyle
        {
            Bold = bold,
            Italic = italic,
            Underline = underline,
            FontName = name,
            FontSize = size,
            Color = color,
        };
    }

    private static bool IsEmptyRunStyle(RichTextStyle s) =>
        s.Bold is null && s.Italic is null && s.Underline is null
        && s.FontName is null && s.FontSize is null && s.Color is null;

    // ---- Helpers ------------------------------------------------------------

    private static bool IsEmpty(CellStyle s) =>
        s.Bold is null && s.Italic is null && s.Underline is null
        && s.FontName is null && s.FontSize is null && s.FontColor is null
        && s.Background is null && s.BackgroundTheme is null
        && s.NumberFormat is null
        && s.HorizontalAlignment is null && s.VerticalAlignment is null
        && s.WrapText is null && s.Borders is null;

    private static bool NeedsFont(CellStyle s) =>
        s.Bold is not null || s.Italic is not null || s.Underline is not null
        || s.FontName is not null || s.FontSize is not null || s.FontColor is not null;

    private static TColor ToColor<TColor>(Color c) where TColor : S.ColorType, new()
        => new() { Rgb = ArgbHex(c) };

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

    private readonly record struct FontKey(string? Name, double? Size, bool Bold, bool Italic, UnderlineStyle Underline, Color? Color)
    {
        internal static FontKey From(CellStyle s) => new(
            s.FontName, s.FontSize, s.Bold ?? false, s.Italic ?? false,
            s.Underline ?? UnderlineStyle.None, s.FontColor);
    }

    private readonly record struct FillKey(Color? Rgb, ThemeColor? Theme)
    {
        internal static FillKey OfRgb(Color c) => new(c, null);
        internal static FillKey OfTheme(ThemeColor t) => new(null, t);
    }
}
