// XssfSheet — image embedding surface per decision I-52.
// Core class structure is in XssfSheet.cs.

using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed partial class XssfSheet
{
    public IPicture AddPicture(string a1Cell, byte[] data, ImageFormat format)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Cell);
        ArgumentNullException.ThrowIfNull(data);

        var (row1, col1) = CellAddress.Parse(a1Cell);
        var npoiPictureType = ToNpoiPictureType(format);
        int pictureIdx = _workbook.Underlying.AddPicture(data, (int)npoiPictureType);

        var drawing = (XSSFDrawing)_underlying.CreateDrawingPatriarch();
        // 0-based row/col; the (0,0,0,0) prefix is per-cell pixel offset
        // (no offset — anchor exactly at the top-left of a1Cell).
        var anchor = new XSSFClientAnchor(0, 0, 0, 0, col1 - 1, row1 - 1, col1 - 1, row1 - 1);
        var picture = (XSSFPicture)drawing.CreatePicture(anchor, pictureIdx);
        // Render at natural pixel size (NPOI sets the to-cell using
        // image dimensions; without Resize the picture would be a
        // single-cell smear).
        picture.Resize();
        return new XssfPicture(_workbook, this, picture, format);
    }

    public IPicture AddPicture(string a1Cell, byte[] data)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Cell);
        ArgumentNullException.ThrowIfNull(data);
        return AddPicture(a1Cell, data, DetectImageFormat(data));
    }

    public IPicture AddPicture(string startCell, string endCell, byte[] data, ImageFormat format)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(startCell);
        ArgumentNullException.ThrowIfNull(endCell);
        ArgumentNullException.ThrowIfNull(data);

        var (r1, c1) = CellAddress.Parse(startCell);
        var (r2, c2) = CellAddress.Parse(endCell);
        var npoiPictureType = ToNpoiPictureType(format);
        int pictureIdx = _workbook.Underlying.AddPicture(data, (int)npoiPictureType);

        var drawing = (XSSFDrawing)_underlying.CreateDrawingPatriarch();
        var anchor = new XSSFClientAnchor(0, 0, 0, 0, c1 - 1, r1 - 1, c2, r2);
        var picture = (XSSFPicture)drawing.CreatePicture(anchor, pictureIdx);
        return new XssfPicture(_workbook, this, picture, format);
    }

    public IPicture AddPicture(string startCell, string endCell, byte[] data)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(startCell);
        ArgumentNullException.ThrowIfNull(endCell);
        ArgumentNullException.ThrowIfNull(data);
        return AddPicture(startCell, endCell, data, DetectImageFormat(data));
    }

    private static PictureType ToNpoiPictureType(ImageFormat f) => f switch
    {
        ImageFormat.Png => PictureType.PNG,
        ImageFormat.Jpeg => PictureType.JPEG,
        _ => throw new ArgumentOutOfRangeException(nameof(f), f, "unsupported ImageFormat"),
    };

    private static ImageFormat DetectImageFormat(byte[] data)
    {
        // PNG: 89 50 4E 47 0D 0A 1A 0A
        if (data.Length >= 8
            && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47
            && data[4] == 0x0D && data[5] == 0x0A && data[6] == 0x1A && data[7] == 0x0A)
        {
            return ImageFormat.Png;
        }
        // JPEG: FF D8 FF
        if (data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF)
        {
            return ImageFormat.Jpeg;
        }
        throw new UnsupportedImageFormatException(
            "The supplied bytes are not a recognized PNG (magic 89 50 4E 47 ...) " +
            "or JPEG (magic FF D8 FF ...). Pass an explicit ImageFormat if the " +
            "format is known and supported.");
    }
}
