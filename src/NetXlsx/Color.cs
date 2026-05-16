// Owned color value type per design decision #29 — no
// System.Drawing.Common dependency (Windows-only since .NET 6).

using System;
using System.Globalization;

namespace NetXlsx;

/// <summary>
/// An immutable ARGB color value. Equality is by ARGB (decision I-23):
/// two colors with identical channel bytes hash and compare equal,
/// regardless of construction site.
/// </summary>
public readonly record struct Color(byte A, byte R, byte G, byte B)
{
    /// <summary>Construct a fully-opaque color from RGB channels.</summary>
    public static Color FromRgb(byte r, byte g, byte b) => new(0xFF, r, g, b);

    /// <summary>Construct a color from explicit ARGB channels.</summary>
    public static Color FromArgb(byte a, byte r, byte g, byte b) => new(a, r, g, b);

    /// <summary>
    /// Parses a hex color string. Accepts <c>"#RRGGBB"</c> and
    /// <c>"#AARRGGBB"</c> forms (the leading <c>#</c> is optional).
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="hex"/> is null.</exception>
    /// <exception cref="FormatException">The string is not in an accepted form.</exception>
    public static Color FromHex(string hex)
    {
        ArgumentNullException.ThrowIfNull(hex);
        var span = hex.AsSpan();
        if (!span.IsEmpty && span[0] == '#') span = span.Slice(1);
        return span.Length switch
        {
            6 => new Color(
                0xFF,
                byte.Parse(span.Slice(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(span.Slice(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(span.Slice(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture)),
            8 => new Color(
                byte.Parse(span.Slice(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(span.Slice(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(span.Slice(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(span.Slice(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture)),
            _ => throw new FormatException(
                $"Hex color must be in the form #RRGGBB or #AARRGGBB; got '{hex}'."),
        };
    }

    /// <summary>Returns the canonical hex representation <c>#AARRGGBB</c>.</summary>
    public string ToHex() => string.Create(9, this, static (span, c) =>
    {
        span[0] = '#';
        WriteByte(span, 1, c.A);
        WriteByte(span, 3, c.R);
        WriteByte(span, 5, c.G);
        WriteByte(span, 7, c.B);
    });

    private static void WriteByte(Span<char> span, int offset, byte value)
    {
        const string hex = "0123456789ABCDEF";
        span[offset] = hex[value >> 4];
        span[offset + 1] = hex[value & 0x0F];
    }

    // Common presets (decision #29 — curated set; full theme/palette
    // arrives with the styling-completion slice).
    /// <summary>Black (#FF000000).</summary>
    public static Color Black => new(0xFF, 0, 0, 0);
    /// <summary>White (#FFFFFFFF).</summary>
    public static Color White => new(0xFF, 255, 255, 255);
    /// <summary>Red (#FFFF0000).</summary>
    public static Color Red => new(0xFF, 255, 0, 0);
    /// <summary>Green (#FF008000).</summary>
    public static Color Green => new(0xFF, 0, 128, 0);
    /// <summary>Blue (#FF0000FF).</summary>
    public static Color Blue => new(0xFF, 0, 0, 255);
    /// <summary>Yellow (#FFFFFF00).</summary>
    public static Color Yellow => new(0xFF, 255, 255, 0);
    /// <summary>Light gray (#FFD3D3D3) — common header fill.</summary>
    public static Color LightGray => new(0xFF, 211, 211, 211);
    /// <summary>Gray (#FF808080).</summary>
    public static Color Gray => new(0xFF, 128, 128, 128);
}
