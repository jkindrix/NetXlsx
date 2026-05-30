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
    /// <summary>PNG (lossless, alpha supported). NPOI <c>PictureType.PNG</c>.</summary>
    Png,
    /// <summary>JPEG (lossy, no alpha). NPOI <c>PictureType.JPEG</c>.</summary>
    Jpeg,
}

/// <summary>
/// A picture anchored on an <see cref="ISheet"/>. Constructed via
/// <see cref="ISheet.AddPicture(string, byte[], ImageFormat)"/>. v1.1
/// exposes the format and sheet — for resizing, anchor manipulation,
/// or alt-text, reach through <see cref="Underlying"/>.
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
    /// Escape hatch — direct access to the underlying NPOI
    /// <c>XSSFPicture</c>. Same contract as
    /// <see cref="IWorkbook.Underlying"/>.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFPicture Underlying { get; }
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
