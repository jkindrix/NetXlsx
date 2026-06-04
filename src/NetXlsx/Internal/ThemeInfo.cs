// Parsed view of an OOXML theme1.xml, used by IWorkbook.ResolveThemeColor
// and GetThemeLineWidthEmu (decision I-81). Built lazily from the
// workbook's theme-part bytes, then cached on the workbook wrapper.

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;

namespace NetXlsx;

internal sealed class ThemeInfo
{
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>Empty (no theme) instance for workbooks without a theme part.</summary>
    public static readonly ThemeInfo Empty = new();

    // Scheme color name -> "RRGGBB" hex. Aliases (tx1/bg1/tx2/bg2) are
    // handled at lookup time, not stored as duplicate entries.
    private readonly Dictionary<string, string> _scheme = new(StringComparer.OrdinalIgnoreCase);

    // Line-style widths in EMU, indexed 0-based (caller passes 1-based and
    // we subtract).
    private readonly List<int> _lineWidthsEmu = new();

    // OOXML cell-color theme index encoding (matches ThemeColor.Index):
    //   0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2,
    //   4..9 = accent1..6, 10 = hlink, 11 = folHlink.
    private static readonly string[] IndexToName =
    {
        "lt1", "dk1", "lt2", "dk2",
        "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
        "hlink", "folHlink",
    };

    public static ThemeInfo Parse(byte[]? xml)
    {
        if (xml is null || xml.Length == 0) return Empty;

        var info = new ThemeInfo();
        try
        {
            var doc = XDocument.Load(new MemoryStream(xml));
            var elements = doc.Root?.Element(A + "themeElements");
            if (elements is null) return info;

            var clrScheme = elements.Element(A + "clrScheme");
            if (clrScheme is not null)
            {
                foreach (var slot in clrScheme.Elements())
                {
                    string name = slot.Name.LocalName;
                    string? hex = ReadColor(slot);
                    if (hex is not null) info._scheme[name] = hex;
                }
            }

            var lnStyleLst = elements.Element(A + "fmtScheme")?.Element(A + "lnStyleLst");
            if (lnStyleLst is not null)
            {
                foreach (var ln in lnStyleLst.Elements(A + "ln"))
                {
                    info._lineWidthsEmu.Add(
                        int.TryParse((string?)ln.Attribute("w"), out int w) ? w : 0);
                }
            }
        }
        catch
        {
            // Best-effort: an unparseable theme yields an empty info.
        }
        return info;
    }

    /// <summary>Resolves a scheme color name (e.g. "dk1", "tx1") to RGB.</summary>
    public Color? ResolveByName(string? name, double tint)
    {
        if (string.IsNullOrEmpty(name)) return null;

        string key = name switch
        {
            "tx1" => "dk1",
            "bg1" => "lt1",
            "tx2" => "dk2",
            "bg2" => "lt2",
            _ => name,
        };
        if (!_scheme.TryGetValue(key, out var hex)) return null;

        return ApplyTint(HexToColor(hex), tint);
    }

    /// <summary>Resolves a 0-based cell-color theme index to RGB.</summary>
    public Color? ResolveByIndex(int index, double tint)
    {
        if (index < 0 || index >= IndexToName.Length) return null;
        return ResolveByName(IndexToName[index], tint);
    }

    /// <summary>EMU width for a 1-based lnRef idx; null if out of range.</summary>
    public int? LineWidthEmu(int oneBasedIdx)
    {
        int i = oneBasedIdx - 1;
        if (i < 0 || i >= _lineWidthsEmu.Count) return null;
        int w = _lineWidthsEmu[i];
        return w > 0 ? w : null;
    }

    private static string? ReadColor(XElement slot)
    {
        var srgb = slot.Element(A + "srgbClr");
        if (srgb is not null) return ((string?)srgb.Attribute("val"))?.ToUpperInvariant();

        var sys = slot.Element(A + "sysClr");
        if (sys is not null) return ((string?)sys.Attribute("lastClr"))?.ToUpperInvariant();

        return null;
    }

    private static Color HexToColor(string hex)
    {
        // 6-char "RRGGBB". Defensive: pad/truncate if needed.
        if (hex.Length < 6) hex = hex.PadRight(6, '0');
        var inv = System.Globalization.CultureInfo.InvariantCulture;
        byte r = byte.Parse(hex.AsSpan(0, 2), System.Globalization.NumberStyles.HexNumber, inv);
        byte g = byte.Parse(hex.AsSpan(2, 2), System.Globalization.NumberStyles.HexNumber, inv);
        byte b = byte.Parse(hex.AsSpan(4, 2), System.Globalization.NumberStyles.HexNumber, inv);
        return Color.FromRgb(r, g, b);
    }

    /// <summary>
    /// Applies Excel's theme-color tint algorithm to an RGB color.
    /// Negative tint darkens, positive lightens; 0 leaves the color
    /// unchanged. Tints are clamped to [-1, +1].
    /// </summary>
    internal static Color ApplyTint(Color baseColor, double tint)
    {
        if (tint == 0) return baseColor;
        if (tint < -1) tint = -1;
        if (tint > 1) tint = 1;

        // RGB → HLS, adjust L per Excel's formula, then HLS → RGB.
        RgbToHls(baseColor.R, baseColor.G, baseColor.B, out double h, out double l, out double s);
        l = tint < 0
            ? l * (1.0 + tint)
            : l * (1.0 - tint) + (1.0 - (1.0 - tint));
        if (l < 0) l = 0; else if (l > 1) l = 1;
        HlsToRgb(h, l, s, out byte r, out byte g, out byte b);
        return Color.FromRgb(r, g, b);
    }

    private static void RgbToHls(byte rByte, byte gByte, byte bByte, out double h, out double l, out double s)
    {
        double r = rByte / 255.0, g = gByte / 255.0, b = bByte / 255.0;
        double max = Math.Max(r, Math.Max(g, b));
        double min = Math.Min(r, Math.Min(g, b));
        l = (max + min) / 2.0;

        if (max == min) { h = 0; s = 0; return; }

        double d = max - min;
        s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);
        if (max == r) h = (g - b) / d + (g < b ? 6.0 : 0.0);
        else if (max == g) h = (b - r) / d + 2.0;
        else h = (r - g) / d + 4.0;
        h /= 6.0;
    }

    private static void HlsToRgb(double h, double l, double s, out byte r, out byte g, out byte b)
    {
        if (s == 0)
        {
            byte v = (byte)Math.Round(l * 255.0);
            r = g = b = v;
            return;
        }
        double q = l < 0.5 ? l * (1.0 + s) : l + s - l * s;
        double p = 2.0 * l - q;
        r = (byte)Math.Round(HueToRgb(p, q, h + 1.0 / 3.0) * 255.0);
        g = (byte)Math.Round(HueToRgb(p, q, h) * 255.0);
        b = (byte)Math.Round(HueToRgb(p, q, h - 1.0 / 3.0) * 255.0);
    }

    private static double HueToRgb(double p, double q, double t)
    {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6.0) return p + (q - p) * 6.0 * t;
        if (t < 1.0 / 2.0) return q;
        if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
        return p;
    }
}
