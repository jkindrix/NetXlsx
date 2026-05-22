// Per-workbook style cache: maps CellStyle value records to NPOI
// ICellStyle instances, deduplicated by CellStyle's structural equality
// (decision #4 / spike 1). This is the v0.4.x replacement for the
// targeted date/time cache (S29) — date/time defaults now flow through
// the same pool as user-supplied styles.

using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SsHAlign = NPOI.SS.UserModel.HorizontalAlignment;
using SsVAlign = NPOI.SS.UserModel.VerticalAlignment;
using SsBorderStyle = NPOI.SS.UserModel.BorderStyle;
using SsFontUnderlineType = NPOI.SS.UserModel.FontUnderlineType;

namespace NetXlsx;

internal sealed class CellStylePool
{
    private readonly XSSFWorkbook _wb;
    private readonly Dictionary<CellStyle, ICellStyle> _pool = new();
    private readonly Dictionary<FontKey, IFont> _fontPool = new();

    // Dedup counters per decision I-61. Read-only operational
    // visibility — increments under the workbook's own mutation
    // protection (the pool is only touched from inside EnterMutation
    // scopes), so plain int suffices.
    internal int StyleHitCount;
    internal int StyleMissCount;
    internal int FontHitCount;
    internal int FontMissCount;
    internal int UniqueStyles => _pool.Count;
    internal int UniqueFonts => _fontPool.Count;

    public CellStylePool(XSSFWorkbook wb) { _wb = wb; }

    /// <summary>
    /// Returns the NPOI <see cref="ICellStyle"/> backing <paramref name="style"/>,
    /// allocating one only if no structurally-equal style is already in
    /// the pool.
    /// </summary>
    public ICellStyle GetOrCreate(CellStyle style)
    {
        if (_pool.TryGetValue(style, out var existing))
        {
            StyleHitCount++;
            return existing;
        }
        StyleMissCount++;
        var npoiStyle = AllocateNpoiStyle(style);
        _pool[style] = npoiStyle;
        return npoiStyle;
    }

    /// <summary>
    /// Reads an existing NPOI cell style back to a <see cref="CellStyle"/>
    /// value record. Lossy for properties NetXlsx does not yet model
    /// (e.g., reading-direction, indent); those properties round-trip
    /// through <c>.Underlying</c> rather than this pool.
    /// </summary>
    public static CellStyle ReadFromNpoi(ICellStyle ns)
    {
        var font = ns.GetFont(null);

        return new CellStyle
        {
            Bold = font.IsBold ? true : null,
            Italic = font.IsItalic ? true : null,
            Underline = MapUnderlineFromNpoi(font.Underline),
            FontName = font.FontName,
            FontSize = font.FontHeightInPoints > 0 ? font.FontHeightInPoints : null,
            FontColor = MapXssfColorToColor((font as XSSFFont)?.GetXSSFColor()),
            Background = MapXssfColorToColor((ns as XSSFCellStyle)?.FillForegroundColorColor as XSSFColor),
            NumberFormat = string.IsNullOrEmpty(ns.GetDataFormatString()) ? null : ns.GetDataFormatString(),
            HorizontalAlignment = MapHAlignFromNpoi(ns.Alignment),
            VerticalAlignment = MapVAlignFromNpoi(ns.VerticalAlignment),
            WrapText = ns.WrapText ? true : null,
            Borders = ReadBordersFromNpoi(ns),
        };
    }

    private static CellBorders? ReadBordersFromNpoi(ICellStyle ns)
    {
        var top = MapBorderStyleFromNpoi(ns.BorderTop);
        var right = MapBorderStyleFromNpoi(ns.BorderRight);
        var bottom = MapBorderStyleFromNpoi(ns.BorderBottom);
        var left = MapBorderStyleFromNpoi(ns.BorderLeft);
        if (top is null && right is null && bottom is null && left is null) return null;
        return new CellBorders(top, null, right, null, bottom, null, left, null);
    }

    private ICellStyle AllocateNpoiStyle(CellStyle style)
    {
        var s = _wb.CreateCellStyle();

        // Font — only allocate one if any font property is set.
        if (NeedsFont(style))
        {
            var font = GetOrCreateFont(new FontKey(
                style.FontName,
                style.FontSize,
                style.Bold ?? false,
                style.Italic ?? false,
                style.Underline ?? UnderlineStyle.None,
                style.FontColor));
            s.SetFont(font);
        }

        // Number format.
        if (!string.IsNullOrEmpty(style.NumberFormat))
        {
            var dataFormat = _wb.CreateDataFormat();
            s.DataFormat = dataFormat.GetFormat(style.NumberFormat);
        }

        // Alignment.
        if (style.HorizontalAlignment is { } h) s.Alignment = MapHAlignToNpoi(h);
        if (style.VerticalAlignment is { } v) s.VerticalAlignment = MapVAlignToNpoi(v);

        // Wrap text.
        if (style.WrapText == true) s.WrapText = true;

        // Background fill.
        if (style.Background is { } bg)
        {
            var xs = (XSSFCellStyle)s;
            xs.SetFillForegroundColor(ToXssfColor(bg));
            s.FillPattern = FillPattern.SolidForeground;
        }

        // Borders.
        if (style.Borders is { } b)
        {
            if (b.Top is { } topStyle)       { s.BorderTop = MapBorderStyleToNpoi(topStyle); }
            if (b.Right is { } rightStyle)   { s.BorderRight = MapBorderStyleToNpoi(rightStyle); }
            if (b.Bottom is { } bottomStyle) { s.BorderBottom = MapBorderStyleToNpoi(bottomStyle); }
            if (b.Left is { } leftStyle)     { s.BorderLeft = MapBorderStyleToNpoi(leftStyle); }

            var xs = (XSSFCellStyle)s;
            if (b.TopColor is { } tc)        xs.SetTopBorderColor(ToXssfColor(tc));
            if (b.RightColor is { } rc)      xs.SetRightBorderColor(ToXssfColor(rc));
            if (b.BottomColor is { } bc)     xs.SetBottomBorderColor(ToXssfColor(bc));
            if (b.LeftColor is { } lc)       xs.SetLeftBorderColor(ToXssfColor(lc));
        }

        return s;
    }

    /// <summary>
    /// Returns the pooled NPOI <see cref="IFont"/> matching the given
    /// rich-text run style. Reuses the same font pool as the cell-style
    /// allocator (decision I-50) — runs with identical font properties
    /// share one <c>IFont</c> across the workbook.
    /// </summary>
    internal IFont GetOrCreateRunFont(RichTextStyle style)
    {
        var key = new FontKey(
            style.FontName,
            style.FontSize,
            style.Bold ?? false,
            style.Italic ?? false,
            style.Underline ?? UnderlineStyle.None,
            style.Color);
        return GetOrCreateFont(key);
    }

    /// <summary>
    /// Reads an NPOI <see cref="IFont"/> back to a <see cref="RichTextStyle"/>.
    /// Mirrors <see cref="ReadFromNpoi"/>'s font axes; lossy for properties
    /// the run model does not cover.
    /// </summary>
    internal static RichTextStyle ReadRunStyleFromFont(IFont font) => new()
    {
        Bold = font.IsBold ? true : null,
        Italic = font.IsItalic ? true : null,
        Underline = MapUnderlineFromNpoi(font.Underline),
        FontName = font.FontName,
        FontSize = font.FontHeightInPoints > 0 ? font.FontHeightInPoints : null,
        Color = MapXssfColorToColor((font as XSSFFont)?.GetXSSFColor()),
    };

    private static bool NeedsFont(CellStyle s) =>
        s.Bold is not null || s.Italic is not null || s.Underline is not null
        || s.FontName is not null || s.FontSize is not null || s.FontColor is not null;

    private IFont GetOrCreateFont(FontKey key)
    {
        if (_fontPool.TryGetValue(key, out var existing))
        {
            FontHitCount++;
            return existing;
        }
        FontMissCount++;
        var f = _wb.CreateFont();
        if (key.Bold) f.IsBold = true;
        if (key.Italic) f.IsItalic = true;
        if (key.Underline != UnderlineStyle.None) f.Underline = MapUnderlineToNpoi(key.Underline);
        if (!string.IsNullOrEmpty(key.Name)) f.FontName = key.Name;
        if (key.Size is { } sz) f.FontHeightInPoints = (short)sz;
        if (key.Color is { } c && f is XSSFFont xf) xf.SetColor(ToXssfColor(c));
        _fontPool[key] = f;
        return f;
    }

    private static XSSFColor ToXssfColor(Color c)
    {
        // The byte[]-only XSSFColor constructor was removed in NPOI 2.7.4+;
        // construct from a CT_Color XML bean instead — that ctor has been
        // present since the 2.5.x line and is stable across all NPOI 2.x.
        var ct = new NPOI.OpenXmlFormats.Spreadsheet.CT_Color
        {
            rgb = new byte[] { 0xFF, c.R, c.G, c.B },  // ARGB; alpha=FF (opaque)
        };
        return new XSSFColor(ct);
    }

    private static Color? MapXssfColorToColor(XSSFColor? xc)
    {
        if (xc is null) return null;
        var rgb = xc.RGB;
        if (rgb is null || rgb.Length < 3) return null;
        return Color.FromRgb(rgb[0], rgb[1], rgb[2]);
    }

    private static SsHAlign MapHAlignToNpoi(HAlign h) => h switch
    {
        HAlign.General => SsHAlign.General,
        HAlign.Left => SsHAlign.Left,
        HAlign.Center => SsHAlign.Center,
        HAlign.Right => SsHAlign.Right,
        HAlign.Fill => SsHAlign.Fill,
        HAlign.Justify => SsHAlign.Justify,
        _ => SsHAlign.General,
    };

    private static HAlign? MapHAlignFromNpoi(SsHAlign h) => h switch
    {
        SsHAlign.General => null,   // null means "inherit / default"
        SsHAlign.Left => HAlign.Left,
        SsHAlign.Center => HAlign.Center,
        SsHAlign.Right => HAlign.Right,
        SsHAlign.Fill => HAlign.Fill,
        SsHAlign.Justify => HAlign.Justify,
        _ => null,
    };

    private static SsVAlign MapVAlignToNpoi(VAlign v) => v switch
    {
        VAlign.Top => SsVAlign.Top,
        VAlign.Center => SsVAlign.Center,
        VAlign.Bottom => SsVAlign.Bottom,
        VAlign.Justify => SsVAlign.Justify,
        _ => SsVAlign.Bottom,
    };

    private static VAlign? MapVAlignFromNpoi(SsVAlign v) => v switch
    {
        SsVAlign.Bottom => null,   // Excel's default — treat as "no explicit value"
        SsVAlign.Top => VAlign.Top,
        SsVAlign.Center => VAlign.Center,
        SsVAlign.Justify => VAlign.Justify,
        _ => null,
    };

    private static SsBorderStyle MapBorderStyleToNpoi(BorderStyle b) => b switch
    {
        BorderStyle.None => SsBorderStyle.None,
        BorderStyle.Thin => SsBorderStyle.Thin,
        BorderStyle.Medium => SsBorderStyle.Medium,
        BorderStyle.Thick => SsBorderStyle.Thick,
        BorderStyle.Double => SsBorderStyle.Double,
        BorderStyle.Dashed => SsBorderStyle.Dashed,
        BorderStyle.Dotted => SsBorderStyle.Dotted,
        _ => SsBorderStyle.None,
    };

    private static BorderStyle? MapBorderStyleFromNpoi(SsBorderStyle b) => b switch
    {
        SsBorderStyle.None => null,
        SsBorderStyle.Thin => BorderStyle.Thin,
        SsBorderStyle.Medium => BorderStyle.Medium,
        SsBorderStyle.Thick => BorderStyle.Thick,
        SsBorderStyle.Double => BorderStyle.Double,
        SsBorderStyle.Dashed => BorderStyle.Dashed,
        SsBorderStyle.Dotted => BorderStyle.Dotted,
        _ => BorderStyle.Thin,   // unmapped NPOI styles fall through as Thin
    };

    private static SsFontUnderlineType MapUnderlineToNpoi(UnderlineStyle u) => u switch
    {
        UnderlineStyle.None => SsFontUnderlineType.None,
        UnderlineStyle.Single => SsFontUnderlineType.Single,
        UnderlineStyle.Double => SsFontUnderlineType.Double,
        UnderlineStyle.SingleAccounting => SsFontUnderlineType.SingleAccounting,
        UnderlineStyle.DoubleAccounting => SsFontUnderlineType.DoubleAccounting,
        _ => SsFontUnderlineType.None,
    };

    private static UnderlineStyle? MapUnderlineFromNpoi(SsFontUnderlineType u) => u switch
    {
        SsFontUnderlineType.None => null,
        SsFontUnderlineType.Single => UnderlineStyle.Single,
        SsFontUnderlineType.Double => UnderlineStyle.Double,
        SsFontUnderlineType.SingleAccounting => UnderlineStyle.SingleAccounting,
        SsFontUnderlineType.DoubleAccounting => UnderlineStyle.DoubleAccounting,
        _ => null,
    };

    private sealed record FontKey(string? Name, double? Size, bool Bold, bool Italic, UnderlineStyle Underline, Color? Color);
}
