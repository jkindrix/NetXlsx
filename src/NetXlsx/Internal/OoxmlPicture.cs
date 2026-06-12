// I-82 engine swap — Open XML SDK-backed IPicture (drawings slice, pictures).
//
// A snapshot of one anchored picture on the SDK engine. Created by
// OoxmlSheet.AddPicture and surfaced (freshly per call) by OoxmlSheet.Pictures.
// The anchor geometry mirrors the NPOI engine's XssfPicture contract exactly so
// the eventual cutover is behavior-preserving (decision I-81):
//
//   FromCell / ToCell  A1 corners. For a one-cell anchor (xdr:oneCellAnchor) there
//                      is no distinct end cell, so ToCell == FromCell and
//                      Dx2 == Dy2 == 0 — the rendered size lives in <xdr:ext>, which
//                      this read surface does not expose (NPOI's IPicture does not
//                      either). Two-cell anchors carry both corners + all four EMU
//                      offsets (lesson #5: each image keeps its own dx/dy; the end
//                      cell is exclusive, the same convention CellAddress round-trips).
//   Data               the raw embedded image bytes.
//
// The escape hatch (Underlying, #32 / I-82) hands out the live xdr:pic element
// the snapshot was created from / parsed out of.

using System;
using A = DocumentFormat.OpenXml.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlPicture : IPicture
{
    private readonly OoxmlWorkbook _workbook;
    private readonly OoxmlSheet _sheet;
    private readonly ImageFormat _format;
    private readonly string _fromCell;
    private readonly string _toCell;
    private readonly int _dx1, _dy1, _dx2, _dy2;
    private readonly byte[] _data;
    private readonly XDR.Picture _pic;

    internal OoxmlPicture(
        OoxmlWorkbook workbook, OoxmlSheet sheet, ImageFormat format,
        string fromCell, string toCell,
        int dx1, int dy1, int dx2, int dy2, byte[] data, XDR.Picture pic)
    {
        _workbook = workbook;
        _sheet = sheet;
        _format = format;
        _fromCell = fromCell;
        _toCell = toCell;
        _dx1 = dx1; _dy1 = dy1; _dx2 = dx2; _dy2 = dy2;
        _data = data;
        _pic = pic;
    }

    public ISheet Sheet { get { _workbook.ThrowIfDisposed(); return _sheet; } }
    public ImageFormat Format { get { _workbook.ThrowIfDisposed(); return _format; } }
    public string FromCell { get { _workbook.ThrowIfDisposed(); return _fromCell; } }
    public string ToCell { get { _workbook.ThrowIfDisposed(); return _toCell; } }
    public int Dx1 { get { _workbook.ThrowIfDisposed(); return _dx1; } }
    public int Dy1 { get { _workbook.ThrowIfDisposed(); return _dy1; } }
    public int Dx2 { get { _workbook.ThrowIfDisposed(); return _dx2; } }
    public int Dy2 { get { _workbook.ThrowIfDisposed(); return _dy2; } }
    public byte[] Data { get { _workbook.ThrowIfDisposed(); return _data; } }

    // Escape hatch (#32 / I-82): the live xdr:pic element. Disposal first.
    public XDR.Picture Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _pic; }
    }

    // ---- Border (decision I-86) ------------------------------------------
    //
    // Reads/writes the live xdr:pic/xdr:spPr/a:ln. Set is a WHOLESALE
    // replacement of <a:ln> (unmodeled line props do not survive a
    // read-modify-write; set-null removes ANY <a:ln> — pinned I-86
    // decisions); get returns null for any border the record cannot
    // represent faithfully (non-solid fills, non-srgb/scheme color models,
    // unmapped scheme names, color-transform children) rather than a
    // silent approximation.

    private const long EmuPerPoint = 12700;          // <a:ln @w> is EMU; the surface is points.
    private const double MaxLineWidthPoints = 1584;  // ST_LineWidth max 20116800 EMU.

    public PictureBorder? Border
    {
        get
        {
            _workbook.ThrowIfDisposed();
            return ReadBorder();
        }
        set
        {
            _workbook.ThrowIfDisposed();
            if (value is null) { _pic.ShapeProperties?.RemoveAllChildren<A.Outline>(); return; }
            ValidateBorder(value);
            WriteBorder(value);
        }
    }

    // The parameter is named `value` so thrown ArgumentExceptions carry the
    // property-setter convention's parameter name.
    private static void ValidateBorder(PictureBorder value)
    {
        if (value.ThemeColor is null && value.Color is null)
        {
            throw new ArgumentException(
                "A PictureBorder must set Color or ThemeColor (a border needs a line color; " +
                "pass null to IPicture.Border to remove the border instead).",
                nameof(value));
        }
        if (value.ThemeColor is { } theme)
        {
            if (theme.Index is < 0 or > 11)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(value), theme.Index,
                    "PictureBorder.ThemeColor.Index must be in the I-81 slot range 0-11 " +
                    "(0=lt1, 1=dk1, 2=lt2, 3=dk2, 4-9=accent1-6, 10=hlink, 11=folHlink).");
            }
            if (theme.Tint != 0)
            {
                throw new ArgumentException(
                    "PictureBorder.ThemeColor.Tint must be 0 — drawingML line colors carry no " +
                    "cell-style tint axis (I-86); author tint-modulated borders through " +
                    "IPicture.Underlying.",
                    nameof(value));
            }
        }
        if (value.WidthPoints is { } w && (!(w > 0) || w > MaxLineWidthPoints))
        {
            throw new ArgumentOutOfRangeException(
                nameof(value), w,
                $"PictureBorder.WidthPoints must be > 0 and <= {MaxLineWidthPoints} points " +
                "(the ST_LineWidth maximum).");
        }
    }

    private void WriteBorder(PictureBorder border)
    {
        // Theme wins over explicit RGB when both are set (the I-79 precedence
        // rule, verbatim from CellStyle.BackgroundTheme). A theme-indexed
        // border participates in the I-89 lazy default-theme embed — without
        // a theme part the schemeClr below is a consumer lottery (LO rewrote
        // exactly this border to literal white in the R-8 evidence).
        A.SolidFill fill;
        if (border.ThemeColor is { } theme)
        {
            _workbook.EnsureThemePart();
            fill = new A.SolidFill(new A.SchemeColor { Val = SchemeValueForIndex(theme.Index) });
        }
        else
        {
            fill = new A.SolidFill(new A.RgbColorModelHex
            {
                Val = $"{border.Color!.Value.R:X2}{border.Color!.Value.G:X2}{border.Color!.Value.B:X2}",
            });
        }

        var ln = new A.Outline(fill);
        if (border.WidthPoints is { } w)
            ln.Width = (int)Math.Round(w * EmuPerPoint);

        var spPr = GetOrCreateShapeProperties();
        spPr.RemoveAllChildren<A.Outline>();
        InsertLine(spPr, ln);
    }

    private XDR.ShapeProperties GetOrCreateShapeProperties()
    {
        if (_pic.ShapeProperties is { } existing) return existing;

        // CT_Picture is a strict sequence (nvPicPr, blipFill, spPr, style? —
        // probe-confirmed against the SDK's compiled particle, no extLst);
        // spPr is schema-required, so this path only runs against a
        // hand-degenerate element — rebuild it in position rather than failing.
        var spPr = new XDR.ShapeProperties();
        if (_pic.GetFirstChild<XDR.ShapeStyle>() is { } style) _pic.InsertBefore(spPr, style);
        else _pic.AppendChild(spPr);
        return spPr;
    }

    // CT_ShapeProperties is a strict sequence; <a:ln> precedes the effect /
    // 3-D / extension tail. On a picture opened from a real file those
    // siblings can already exist, so position the new <a:ln> before the first
    // of them (the SDK-quirk #8 discipline, applied locally — OoxmlSchemaOrder
    // itself covers only CT_Worksheet/CT_Workbook).
    private static void InsertLine(XDR.ShapeProperties spPr, A.Outline ln)
    {
        foreach (var child in spPr.ChildElements)
        {
            if (child is A.EffectList or A.EffectDag or A.Scene3DType or A.Shape3DType
                or A.ShapePropertiesExtensionList)
            {
                spPr.InsertBefore(ln, child);
                return;
            }
        }
        spPr.AppendChild(ln);
    }

    private PictureBorder? ReadBorder()
    {
        var ln = _pic.ShapeProperties?.GetFirstChild<A.Outline>();
        if (ln is null) return null;

        // Only a solid fill is representable; <a:noFill> (and absent fill),
        // gradients and patterns read as null (pinned I-86 decision — the
        // a:ln w="1" + a:noFill borderless idiom renders identically to no
        // <a:ln> at all).
        var solid = ln.GetFirstChild<A.SolidFill>();
        if (solid is null) return null;

        double? width = ln.Width?.Value is int w ? w / (double)EmuPerPoint : null;

        if (solid.GetFirstChild<A.RgbColorModelHex>() is { } srgb)
        {
            // Color-transform children (alpha, lumMod, ...) are not modeled —
            // reading them as the bare color would misreport the rendering.
            if (srgb.HasChildren) return null;
            return new PictureBorder { Color = ParseSrgbHex(srgb.Val?.Value), WidthPoints = width };
        }
        if (solid.GetFirstChild<A.SchemeColor>() is { } scheme)
        {
            if (scheme.HasChildren) return null;
            if (ThemeInfo.SchemeNameToIndex(scheme.Val?.InnerText) is not int index) return null;
            return new PictureBorder { ThemeColor = new ThemeColor(index), WidthPoints = width };
        }
        // hslClr / sysClr / scrgbClr / prstClr: not representable.
        return null;
    }

    // ST_HexColorRGB is exactly RRGGBB; anything else in an opened file is
    // genuine corruption — fail loud per the I-83/quirk-#13 discipline, never
    // silently substitute a color.
    private static Color ParseSrgbHex(string? val)
    {
        if (val is { Length: 6 })
        {
            try { return Color.FromHex(val); }
            catch (FormatException) { /* fall through to throw below */ }
        }
        throw new MalformedFileException(
            $"Picture border color 'a:srgbClr/@val' must be a 6-digit RRGGBB hex value; got '{val ?? "(missing)"}'.");
    }

    // The I-81 slot map, emission side (read side: ThemeInfo.SchemeNameToIndex).
    private static A.SchemeColorValues SchemeValueForIndex(int index) => index switch
    {
        0 => A.SchemeColorValues.Light1,
        1 => A.SchemeColorValues.Dark1,
        2 => A.SchemeColorValues.Light2,
        3 => A.SchemeColorValues.Dark2,
        4 => A.SchemeColorValues.Accent1,
        5 => A.SchemeColorValues.Accent2,
        6 => A.SchemeColorValues.Accent3,
        7 => A.SchemeColorValues.Accent4,
        8 => A.SchemeColorValues.Accent5,
        9 => A.SchemeColorValues.Accent6,
        10 => A.SchemeColorValues.Hyperlink,
        11 => A.SchemeColorValues.FollowedHyperlink,
        _ => throw new ArgumentOutOfRangeException(nameof(index), index, "theme slot index must be 0-11"),
    };
}
