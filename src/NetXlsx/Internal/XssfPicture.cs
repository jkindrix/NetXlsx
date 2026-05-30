// XssfPicture — internal wrapper around NPOI's XSSFPicture.
// Created via XssfSheet.AddPicture or surfaced via XssfSheet.Pictures.

using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfPicture : IPicture
{
    private readonly XssfWorkbook _workbook;
    private readonly XssfSheet _sheet;
    private readonly XSSFPicture _underlying;
    private readonly ImageFormat _format;

    public XssfPicture(XssfWorkbook workbook, XssfSheet sheet, XSSFPicture underlying, ImageFormat format)
    {
        _workbook = workbook;
        _sheet = sheet;
        _underlying = underlying;
        _format = format;
    }

    /// <summary>
    /// Wraps an existing NPOI picture (read path, decision I-81) — derives
    /// the format from its MIME type, falling back to PNG for the rare
    /// unrecognized case (the underlying bytes are still preserved via
    /// <see cref="Data"/>).
    /// </summary>
    internal static XssfPicture FromExisting(XssfWorkbook workbook, XssfSheet sheet, XSSFPicture picture)
    {
        var fmt = picture.PictureData.MimeType switch
        {
            "image/png" => ImageFormat.Png,
            "image/jpeg" or "image/jpg" => ImageFormat.Jpeg,
            _ => ImageFormat.Png,
        };
        return new XssfPicture(workbook, sheet, picture, fmt);
    }

    public ISheet Sheet
    {
        get { _workbook.ThrowIfDisposed(); return _sheet; }
    }

    public ImageFormat Format
    {
        get { _workbook.ThrowIfDisposed(); return _format; }
    }

    public string FromCell
    {
        get { _workbook.ThrowIfDisposed(); var a = Anchor; return CellAddress.Format(a.Row1 + 1, a.Col1 + 1); }
    }

    public string ToCell
    {
        get { _workbook.ThrowIfDisposed(); var a = Anchor; return CellAddress.Format(a.Row2 + 1, a.Col2 + 1); }
    }

    public int Dx1 { get { _workbook.ThrowIfDisposed(); return Anchor.Dx1; } }
    public int Dy1 { get { _workbook.ThrowIfDisposed(); return Anchor.Dy1; } }
    public int Dx2 { get { _workbook.ThrowIfDisposed(); return Anchor.Dx2; } }
    public int Dy2 { get { _workbook.ThrowIfDisposed(); return Anchor.Dy2; } }

    public byte[] Data
    {
        get { _workbook.ThrowIfDisposed(); return _underlying.PictureData.Data; }
    }

    public XSSFPicture Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _underlying; }
    }

    private XSSFClientAnchor Anchor => (XSSFClientAnchor)_underlying.ClientAnchor;
}
