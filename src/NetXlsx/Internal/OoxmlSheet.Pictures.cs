// I-82 engine swap — Open XML SDK-backed pictures (drawings slice).
//
// OOXML stores a sheet's images in a DrawingsPart (xl/drawings/drawingN.xml,
// root xdr:wsDr) referenced by a worksheet <drawing r:id> child; the image bytes
// live in an ImagePart referenced by each picture's blip fill (a:blip/@r:embed).
// The worksheet <drawing> sits near the end of CT_Worksheet's child sequence
// (rank "drawing"), so it is routed through OoxmlSchemaOrder.GetOrInsert — correct
// even on an opened file that already carries intervening siblings (SDK-quirk #8).
//
// Anchor strategy (mirrors the public contract the NPOI engine established, I-52/I-81):
//   - Single-cell overloads anchor at the top-left of one cell at the image's
//     NATURAL pixel size via an xdr:oneCellAnchor with an <xdr:ext> in EMU. This is
//     cleaner than the NPOI engine's CreatePicture+Resize() dance (which depends on
//     column-width metrics) and renders identically — exactly the kind of fidelity
//     win the SDK engine unlocks (continuation lesson #5 / the drawings rationale).
//   - Range overloads anchor across two cells via an xdr:twoCellAnchor, preserving
//     per-image EMU offsets dx1/dy1/dx2/dy2 (lesson #5). The end cell is exclusive,
//     the convention CellAddress already round-trips (FromCell/ToCell add 1 back).
//
// CellAddress (1-based) maps to OOXML markers (0-based) by subtracting 1; the read
// surface adds it back, so AddPicture("B3","D6",...) round-trips to FromCell "B3" /
// ToCell "D6", identical to the NPOI engine.

using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    // 96-DPI pixel -> EMU (English Metric Units): 914400 EMU per inch / 96.
    private const long EmuPerPixel = 9525;

    public IPicture AddPicture(string a1Cell, byte[] data, ImageFormat format)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(a1Cell);
        ArgumentNullException.ThrowIfNull(data);

        var (row, col) = CellAddress.Parse(a1Cell);
        var (wPx, hPx) = ImageDimensions(data, format);
        long cx = wPx * EmuPerPixel, cy = hPx * EmuPerPixel;

        var (drawingsPart, root) = GetOrCreateDrawing();
        string embed = EmbedImage(drawingsPart, data, format);

        var pic = BuildPic(root, embed, cx, cy);
        var anchor = new XDR.OneCellAnchor(
            Marker(col - 1, 0, row - 1, 0),
            new XDR.Extent { Cx = cx, Cy = cy },
            pic,
            new XDR.ClientData());
        root.AppendChild(anchor);

        return new OoxmlPicture(_workbook, this, format, a1Cell, a1Cell, 0, 0, 0, 0, data, pic);
    }

    public IPicture AddPicture(string a1Cell, byte[] data)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(a1Cell);
        ArgumentNullException.ThrowIfNull(data);
        return AddPicture(a1Cell, data, DetectImageFormat(data));
    }

    public IPicture AddPicture(string startCell, string endCell, byte[] data, ImageFormat format)
        => AddPicture(startCell, endCell, data, format, 0, 0, 0, 0);

    public IPicture AddPicture(string startCell, string endCell, byte[] data)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(startCell);
        ArgumentNullException.ThrowIfNull(endCell);
        ArgumentNullException.ThrowIfNull(data);
        return AddPicture(startCell, endCell, data, DetectImageFormat(data), 0, 0, 0, 0);
    }

    public IPicture AddPicture(string startCell, string endCell, byte[] data, ImageFormat format,
        int dx1, int dy1, int dx2, int dy2)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(startCell);
        ArgumentNullException.ThrowIfNull(endCell);
        ArgumentNullException.ThrowIfNull(data);

        var (r1, c1) = CellAddress.Parse(startCell);
        var (r2, c2) = CellAddress.Parse(endCell);
        var (wPx, hPx) = ImageDimensions(data, format);
        long cx = wPx * EmuPerPixel, cy = hPx * EmuPerPixel;

        var (drawingsPart, root) = GetOrCreateDrawing();
        string embed = EmbedImage(drawingsPart, data, format);

        var pic = BuildPic(root, embed, cx, cy);
        var anchor = new XDR.TwoCellAnchor(
            Marker(c1 - 1, dx1, r1 - 1, dy1),
            ToMarker(c2 - 1, dx2, r2 - 1, dy2),
            pic,
            new XDR.ClientData());
        root.AppendChild(anchor);

        return new OoxmlPicture(_workbook, this, format,
            CellAddress.Format(r1, c1), CellAddress.Format(r2, c2),
            dx1, dy1, dx2, dy2, data, pic);
    }

    public IReadOnlyList<IPicture> Pictures
    {
        get
        {
            ThrowIfUnusable();
            var result = new List<IPicture>();
            var drawingEl = Worksheet.GetFirstChild<S.Drawing>();
            if (drawingEl?.Id?.Value is not string rid) return result;
            if (_worksheetPart.GetPartById(rid) is not DrawingsPart dp) return result;
            var root = dp.WorksheetDrawing;
            if (root is null) return result;

            foreach (var child in root.ChildElements)
            {
                switch (child)
                {
                    case XDR.TwoCellAnchor twoCell when twoCell.GetFirstChild<XDR.Picture>() is { } p2:
                        result.Add(ReadTwoCell(dp, twoCell, p2));
                        break;
                    case XDR.OneCellAnchor oneCell when oneCell.GetFirstChild<XDR.Picture>() is { } p1:
                        result.Add(ReadOneCell(dp, oneCell, p1));
                        break;
                }
            }
            return result;
        }
    }

    // ---- DrawingsPart / picture construction --------------------------------

    private (DrawingsPart Part, XDR.WorksheetDrawing Root) GetOrCreateDrawing()
    {
        var existing = Worksheet.GetFirstChild<S.Drawing>();
        if (existing?.Id?.Value is string rid
            && _worksheetPart.GetPartById(rid) is DrawingsPart found)
        {
            return (found, found.WorksheetDrawing ??= new XDR.WorksheetDrawing());
        }

        var drawingsPart = _worksheetPart.AddNewPart<DrawingsPart>();
        drawingsPart.WorksheetDrawing = new XDR.WorksheetDrawing();
        string newRid = _worksheetPart.GetIdOfPart(drawingsPart);
        // <drawing> is a strict-sequence worksheet child; place it by schema order
        // so an opened file with later siblings (e.g. <tableParts>) stays valid.
        OoxmlSchemaOrder.GetOrInsert(Worksheet, () => new S.Drawing { Id = newRid });
        return (drawingsPart, drawingsPart.WorksheetDrawing);
    }

    private static string EmbedImage(DrawingsPart drawingsPart, byte[] data, ImageFormat format)
    {
        var imagePart = drawingsPart.AddImagePart(ToImagePartType(format));
        using (var ms = new MemoryStream(data, writable: false))
            imagePart.FeedData(ms);
        return drawingsPart.GetIdOfPart(imagePart);
    }

    private static XDR.Picture BuildPic(XDR.WorksheetDrawing root, string embed, long cx, long cy)
    {
        uint id = NextShapeId(root);
        return new XDR.Picture(
            new XDR.NonVisualPictureProperties(
                new XDR.NonVisualDrawingProperties { Id = id, Name = $"Picture {id}" },
                new XDR.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
            new XDR.BlipFill(
                new A.Blip { Embed = embed },
                new A.Stretch(new A.FillRectangle())),
            new XDR.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
    }

    // cNvPr/@id must be a unique non-zero uint within the drawing.
    private static uint NextShapeId(XDR.WorksheetDrawing root)
    {
        uint next = 1;
        foreach (var p in root.Descendants<XDR.NonVisualDrawingProperties>())
            if (p.Id?.Value is uint id && id >= next) next = id + 1;
        return next;
    }

    private static XDR.FromMarker Marker(int col, int colOff, int row, int rowOff) =>
        new(new XDR.ColumnId(col.ToString(System.Globalization.CultureInfo.InvariantCulture)),
            new XDR.ColumnOffset(colOff.ToString(System.Globalization.CultureInfo.InvariantCulture)),
            new XDR.RowId(row.ToString(System.Globalization.CultureInfo.InvariantCulture)),
            new XDR.RowOffset(rowOff.ToString(System.Globalization.CultureInfo.InvariantCulture)));

    private static XDR.ToMarker ToMarker(int col, int colOff, int row, int rowOff) =>
        new(new XDR.ColumnId(col.ToString(System.Globalization.CultureInfo.InvariantCulture)),
            new XDR.ColumnOffset(colOff.ToString(System.Globalization.CultureInfo.InvariantCulture)),
            new XDR.RowId(row.ToString(System.Globalization.CultureInfo.InvariantCulture)),
            new XDR.RowOffset(rowOff.ToString(System.Globalization.CultureInfo.InvariantCulture)));

    // ---- Read-back ----------------------------------------------------------

    private OoxmlPicture ReadTwoCell(DrawingsPart dp, XDR.TwoCellAnchor anchor, XDR.Picture pic)
    {
        var from = anchor.GetFirstChild<XDR.FromMarker>()!;
        var to = anchor.GetFirstChild<XDR.ToMarker>()!;
        int fc = ParseMarker(from.ColumnId), fco = ParseMarker(from.ColumnOffset);
        int fr = ParseMarker(from.RowId), fro = ParseMarker(from.RowOffset);
        int tc = ParseMarker(to.ColumnId), tco = ParseMarker(to.ColumnOffset);
        int tr = ParseMarker(to.RowId), tro = ParseMarker(to.RowOffset);
        var (data, format) = ReadImage(dp, pic);
        return new OoxmlPicture(_workbook, this, format,
            CellAddress.Format(fr + 1, fc + 1), CellAddress.Format(tr + 1, tc + 1),
            fco, fro, tco, tro, data, pic);
    }

    private OoxmlPicture ReadOneCell(DrawingsPart dp, XDR.OneCellAnchor anchor, XDR.Picture pic)
    {
        var from = anchor.GetFirstChild<XDR.FromMarker>()!;
        int fc = ParseMarker(from.ColumnId), fco = ParseMarker(from.ColumnOffset);
        int fr = ParseMarker(from.RowId), fro = ParseMarker(from.RowOffset);
        var (data, format) = ReadImage(dp, pic);
        // A one-cell anchor has no distinct end cell — surface ToCell == FromCell,
        // Dx2 == Dy2 == 0 (the rendered size lives in <xdr:ext>, which IPicture
        // does not expose).
        string cell = CellAddress.Format(fr + 1, fc + 1);
        return new OoxmlPicture(_workbook, this, format, cell, cell, fco, fro, 0, 0, data, pic);
    }

    private static (byte[] Data, ImageFormat Format) ReadImage(DrawingsPart dp, XDR.Picture pic)
    {
        string? embed = pic.BlipFill?.GetFirstChild<A.Blip>()?.Embed?.Value;
        if (embed is null || dp.GetPartById(embed) is not ImagePart imagePart)
            return (Array.Empty<byte>(), ImageFormat.Png);

        byte[] data;
        using (var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
        using (var ms = new MemoryStream())
        {
            stream.CopyTo(ms);
            data = ms.ToArray();
        }
        var format = imagePart.ContentType switch
        {
            "image/png" => ImageFormat.Png,
            "image/jpeg" or "image/jpg" => ImageFormat.Jpeg,
            _ => ImageFormat.Png,
        };
        return (data, format);
    }

    // ---- Image helpers ------------------------------------------------------

    private static PartTypeInfo ToImagePartType(ImageFormat f) => f switch
    {
        ImageFormat.Png => ImagePartType.Png,
        ImageFormat.Jpeg => ImagePartType.Jpeg,
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

    // Natural pixel dimensions, for sizing the one-cell anchor's <xdr:ext>. Falls
    // back to 1x1 for an unparseable header (the picture still embeds + renders;
    // only its default extent is degenerate, recoverable by passing a range anchor).
    private static (int Width, int Height) ImageDimensions(byte[] data, ImageFormat format) =>
        format switch
        {
            ImageFormat.Png => PngDimensions(data),
            ImageFormat.Jpeg => JpegDimensions(data),
            _ => (1, 1),
        };

    private static (int, int) PngDimensions(byte[] d)
    {
        // IHDR width/height are big-endian uint32 at byte 16 / 20 (8-byte signature
        // + 4-byte length + 4-byte "IHDR").
        if (d.Length < 24) return (1, 1);
        int w = (d[16] << 24) | (d[17] << 16) | (d[18] << 8) | d[19];
        int h = (d[20] << 24) | (d[21] << 16) | (d[22] << 8) | d[23];
        return (w > 0 ? w : 1, h > 0 ? h : 1);
    }

    private static (int, int) JpegDimensions(byte[] d)
    {
        // Walk the marker segments to the first Start-Of-Frame (SOFn), which carries
        // the image height then width as big-endian uint16.
        int pos = 2; // skip SOI (FF D8)
        while (pos + 9 < d.Length)
        {
            if (d[pos] != 0xFF) { pos++; continue; }
            byte marker = d[pos + 1];
            // Standalone markers (no length payload): SOI/EOI, RSTn, TEM.
            if (marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7) || marker == 0x01)
            {
                pos += 2;
                continue;
            }
            int len = (d[pos + 2] << 8) | d[pos + 3];
            // SOF0..SOF15 except DHT (C4), JPG (C8), DAC (CC) carry frame dimensions.
            if (marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC)
            {
                int h = (d[pos + 5] << 8) | d[pos + 6];
                int w = (d[pos + 7] << 8) | d[pos + 8];
                return (w > 0 ? w : 1, h > 0 ? h : 1);
            }
            pos += 2 + len;
        }
        return (1, 1);
    }
}
