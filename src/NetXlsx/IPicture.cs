// Image embedding per design §6.4.2 / decision I-52.
// v1.1 supports PNG and JPEG anchored to a single top-left cell at
// the image's natural pixel size. Multi-cell anchoring and other
// formats (GIF, BMP, TIFF) reach through .Underlying.

namespace NetXlsx;

/// <summary>
/// Image formats supported by <see cref="ISheet.AddPicture(string, byte[], ImageFormat)"/>.
/// v1.1 ships PNG and JPEG only — the two formats Excel reads on
/// every platform without theme-color quirks. Other formats (GIF,
/// BMP, TIFF, EMF, WMF) are reachable through
/// <see cref="ISheet.Underlying"/>.
/// </summary>
public enum ImageFormat
{
    /// <summary>PNG (lossless, alpha supported).</summary>
    Png,
    /// <summary>JPEG (lossy, no alpha).</summary>
    Jpeg,
}

/// <summary>
/// A solid line border around a picture (decision I-86). Set via
/// <see cref="IPicture.Border"/>; written as
/// <c>&lt;a:ln&gt;&lt;a:solidFill&gt;…&lt;/a:solidFill&gt;&lt;/a:ln&gt;</c>
/// on the picture's shape properties.
/// </summary>
public sealed record PictureBorder
{
    /// <summary>Explicit sRGB border color. The alpha channel is ignored
    /// (drawing line colors are emitted as <c>a:srgbClr</c> RGB — the same
    /// convention as <see cref="ISheet.AddShape"/>'s line color).</summary>
    public Color? Color { get; init; }

    /// <summary>
    /// Theme-based border color (the I-81 slot encoding: 0 = lt1, 1 = dk1,
    /// 2 = lt2, 3 = dk2, 4–9 = accent1–6, 10 = hlink, 11 = folHlink).
    /// When set, takes precedence over <see cref="Color"/> (the I-79
    /// precedence rule) and is written as <c>a:schemeClr</c>.
    /// <see cref="NetXlsx.ThemeColor.Tint"/> must be 0 — drawing line
    /// colors carry no cell-style tint axis; tint-modulated borders are
    /// reachable through <see cref="IPicture.Underlying"/>.
    /// </summary>
    public ThemeColor? ThemeColor { get; init; }

    /// <summary>
    /// Line width in points (1 pt = 12700 EMU), or null to omit the
    /// width attribute (Excel renders its default hairline width).
    /// Must be &gt; 0 and ≤ 1584 (the ST_LineWidth maximum).
    /// </summary>
    public double? WidthPoints { get; init; }
}

/// <summary>
/// A picture anchored on an <see cref="ISheet"/>. Constructed via
/// <see cref="ISheet.AddPicture(string, byte[], ImageFormat)"/>. Exposes
/// the format, sheet, anchor geometry, and the solid line border
/// (<see cref="Border"/>, decision I-86) — for resizing, anchor
/// manipulation, or alt-text, reach through <see cref="Underlying"/>.
/// </summary>
public interface IPicture
{
    /// <summary>The owning sheet.</summary>
    ISheet Sheet { get; }

    /// <summary>The picture's stored format.</summary>
    ImageFormat Format { get; }

    /// <summary>
    /// The picture's top-left anchor cell in A1 notation (decision I-81).
    /// For a two-cell anchor this is the <c>from</c> cell.
    /// </summary>
    string FromCell { get; }

    /// <summary>
    /// The picture's bottom-right anchor cell in A1 notation (decision
    /// I-81). For a one-cell anchor this equals <see cref="FromCell"/>.
    /// </summary>
    string ToCell { get; }

    /// <summary>EMU x-offset of the start anchor within <see cref="FromCell"/> (I-81).</summary>
    int Dx1 { get; }
    /// <summary>EMU y-offset of the start anchor within <see cref="FromCell"/> (I-81).</summary>
    int Dy1 { get; }
    /// <summary>EMU x-offset of the end anchor within <see cref="ToCell"/> (I-81).</summary>
    int Dx2 { get; }
    /// <summary>EMU y-offset of the end anchor within <see cref="ToCell"/> (I-81).</summary>
    int Dy2 { get; }

    /// <summary>The raw image bytes (decision I-81).</summary>
    byte[] Data { get; }

    /// <summary>
    /// The picture's solid line border (decision I-86), or null when the
    /// picture has no border (no <c>&lt;a:ln&gt;</c> element, or one whose
    /// fill is <c>a:noFill</c> — both render borderless).
    /// <para>
    /// <b>Set is a wholesale replacement of the <c>&lt;a:ln&gt;</c>
    /// element:</b> line properties this record does not model (dash
    /// style, gradient fill, caps, compound lines, …) do NOT survive a
    /// read-modify-write, and setting null removes ANY existing
    /// <c>&lt;a:ln&gt;</c> including a non-solid one. Full line control
    /// remains available through <see cref="Underlying"/>.
    /// </para>
    /// <para>
    /// Get returns null for borders this record cannot represent
    /// faithfully: non-solid fills, color models other than
    /// <c>a:srgbClr</c>/<c>a:schemeClr</c>, scheme names outside the I-81
    /// slot map (e.g. <c>phClr</c>), and colors carrying transform child
    /// elements (alpha, lumMod, …) — a documented divergence, never a
    /// silent approximation.
    /// </para>
    /// </summary>
    /// <exception cref="System.ArgumentException">On set: neither
    /// <see cref="PictureBorder.Color"/> nor
    /// <see cref="PictureBorder.ThemeColor"/> is set, or the theme color
    /// has a non-zero tint.</exception>
    /// <exception cref="System.ArgumentOutOfRangeException">On set: the
    /// theme index is outside 0–11, or
    /// <see cref="PictureBorder.WidthPoints"/> is not in (0, 1584].</exception>
    PictureBorder? Border { get; set; }

    /// <summary>
    /// Escape hatch — direct access to the underlying Open XML SDK
    /// <see cref="DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture"/>
    /// element (I-82). Same contract as
    /// <see cref="IWorkbook.Underlying"/>.
    /// </summary>
    DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture Underlying { get; }
}

/// <summary>
/// Thrown when <see cref="ISheet.AddPicture(string, byte[])"/> cannot
/// classify the byte buffer as a supported image format. Inspect the
/// data's magic bytes or pass an explicit
/// <see cref="ImageFormat"/> via the 3-arg overload.
/// </summary>
public sealed class UnsupportedImageFormatException : WorkbookException
{
    /// <summary>Constructs the exception with a default message.</summary>
    public UnsupportedImageFormatException()
        : base("The supplied bytes are not a recognized image format (PNG or JPEG expected).") { }

    /// <summary>Constructs the exception with a custom message.</summary>
    public UnsupportedImageFormatException(string message) : base(message) { }

    /// <summary>Constructs the exception with a custom message and inner exception.</summary>
    public UnsupportedImageFormatException(string message, System.Exception innerException)
        : base(message, innerException) { }
}
