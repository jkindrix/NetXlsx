// I-82 engine swap — drawings slice: picture conformance.
//
// Mirrors the NPOI engine's ISheet.AddPicture / ISheet.Pictures contract on the
// Open XML SDK engine (the NetXlsx.Tests references for parity: PictureApiTests,
// RowHeightAndPictureAnchorTests, the Pictures section of ThemeReadAndDrawing
// IterationTests): the five AddPicture overloads, magic-byte auto-detection, the
// validation surface, and the IPicture read-back (FromCell/ToCell/Dx*/Dy*/Data/
// Format). It additionally pins the SDK engine's own behavior that the NPOI engine
// could not express:
//   - single-cell overloads emit an xdr:oneCellAnchor at the image's NATURAL pixel
//     size (no CreatePicture+Resize() column-width dance) — verified via <xdr:ext>;
//   - two-cell overloads emit an xdr:twoCellAnchor preserving per-image EMU offsets
//     (lesson #5), the end cell exclusive (CellAddress round-trips it);
//   - the NPOI escape hatch (IPicture.Underlying) throws NotSupportedException.
// Every fixture round-trips clean through OpenXmlValidationGate where it builds
// real drawing parts.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml.Packaging;
using NetXlsx;
using Xunit;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace NetXlsx.OoxmlEngine.Tests;

public class PictureTests
{
    private const long EmuPerPixel = 9525;

    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-pic-{Guid.NewGuid():N}.xlsx");

    // Known-valid 1×1 transparent PNG.
    private static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

    // Minimal valid JPEG (1×1 grayscale); the leading FF D8 FF is what the
    // magic-byte detector checks, the SOF0 segment carries the 1×1 dimensions.
    private static readonly byte[] TinyJpeg = new byte[]
    {
        0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01,
        0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0xFF, 0xDB, 0x00, 0x43,
        0x00, 0x08, 0x06, 0x06, 0x07, 0x06, 0x05, 0x08, 0x07, 0x07, 0x07, 0x09,
        0x09, 0x08, 0x0A, 0x0C, 0x14, 0x0D, 0x0C, 0x0B, 0x0B, 0x0C, 0x19, 0x12,
        0x13, 0x0F, 0x14, 0x1D, 0x1A, 0x1F, 0x1E, 0x1D, 0x1A, 0x1C, 0x1C, 0x20,
        0x24, 0x2E, 0x27, 0x20, 0x22, 0x2C, 0x23, 0x1C, 0x1C, 0x28, 0x37, 0x29,
        0x2C, 0x30, 0x31, 0x34, 0x34, 0x34, 0x1F, 0x27, 0x39, 0x3D, 0x38, 0x32,
        0x3C, 0x2E, 0x33, 0x34, 0x32, 0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00, 0x01,
        0x00, 0x01, 0x01, 0x01, 0x11, 0x00, 0xFF, 0xC4, 0x00, 0x14, 0x00, 0x01,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0xFF, 0xC4, 0x00, 0x14, 0x10, 0x01, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00,
        0xB2, 0xC0, 0x07, 0xFF, 0xD9,
    };

    // A structurally-valid PNG header carrying an arbitrary IHDR width/height, used
    // to verify natural-size extent computation (the bytes need only round-trip and
    // expose correct dimensions; the image need not be renderable for these checks).
    private static byte[] MakePng(int width, int height)
    {
        var png = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR length + type
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // width(4) + height(4)
            0x08, 0x02, 0x00, 0x00, 0x00,                   // bitdepth/colortype/...
            0x00, 0x00, 0x00, 0x00,                         // (fake CRC)
            0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND length + type
            0xAE, 0x42, 0x60, 0x82,                         // IEND CRC
        };
        png[16] = (byte)(width >> 24); png[17] = (byte)(width >> 16);
        png[18] = (byte)(width >> 8); png[19] = (byte)width;
        png[20] = (byte)(height >> 24); png[21] = (byte)(height >> 16);
        png[22] = (byte)(height >> 8); png[23] = (byte)height;
        return png;
    }

    private static int PictureCount(IWorkbook wb, string sheet)
    {
        var wsp = wb.Underlying.WorkbookPart!.WorksheetParts;
        // The single sheet's drawing part, when present, holds the pictures.
        var dp = wsp.SelectMany(p => p.GetPartsOfType<DrawingsPart>()).FirstOrDefault();
        if (dp?.WorksheetDrawing is null) return 0;
        return dp.WorksheetDrawing.Descendants<XDR.Picture>().Count();
    }

    // ---- Happy path: single-cell overloads ----------------------------------

    [Fact]
    public void AddPicture_With_Explicit_Format_Returns_Picture()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        var pic = sh.AddPicture("B2", OnePixelPng, ImageFormat.Png);
        pic.Sheet.Should().BeSameAs(sh);
        pic.Format.Should().Be(ImageFormat.Png);
        pic.FromCell.Should().Be("B2");
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void AddPicture_AutoDetects_PNG_From_Magic_Bytes()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.AddPicture("A1", OnePixelPng).Format.Should().Be(ImageFormat.Png);
    }

    [Fact]
    public void AddPicture_AutoDetects_JPEG_From_Magic_Bytes()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.AddPicture("A1", TinyJpeg).Format.Should().Be(ImageFormat.Jpeg);
    }

    [Fact]
    public void AddPicture_SingleCell_Uses_A_OneCellAnchor_At_Natural_Size()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.AddPicture("C4", MakePng(10, 20), ImageFormat.Png);

        var dp = wb.Underlying.WorkbookPart!.WorksheetParts
            .SelectMany(p => p.GetPartsOfType<DrawingsPart>()).Single();
        var root = dp.WorksheetDrawing!;
        root.Elements<XDR.OneCellAnchor>().Should().ContainSingle();
        var ext = root.Descendants<XDR.Extent>().Single();
        ext.Cx!.Value.Should().Be(10 * EmuPerPixel);
        ext.Cy!.Value.Should().Be(20 * EmuPerPixel);
        OpenXmlValidationGate.AssertValid(wb);
    }

    // Pins the JPEG SOF dimension walk against a real multi-pixel frame: the only other
    // JPEG fixture (TinyJpeg) is 1×1, indistinguishable from JpegDimensions' (1,1)
    // fallback, so a broken SOF walk would pass every other JPEG test. Distinct width
    // and height (3×2) also catch a transposed height/width read.
    [Fact]
    public void AddPicture_SingleCell_JPEG_Uses_The_Parsed_SOF_Dimensions()
    {
        // SOI (FF D8) + SOF0 (FF C0, len 000B, precision 08, height 0002, width 0003,
        // 1 component) + EOI (FF D9). The bytes need only parse + round-trip, not render.
        var jpeg = new byte[]
        {
            0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00, 0x02, 0x00, 0x03,
            0x01, 0x01, 0x11, 0x00, 0xFF, 0xD9,
        };
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S").AddPicture("A1", jpeg).Format.Should().Be(ImageFormat.Jpeg);

        var ext = wb.Underlying.WorkbookPart!.WorksheetParts
            .SelectMany(p => p.GetPartsOfType<DrawingsPart>()).Single()
            .WorksheetDrawing!.Descendants<XDR.Extent>().Single();
        ext.Cx!.Value.Should().Be(3 * EmuPerPixel);   // width
        ext.Cy!.Value.Should().Be(2 * EmuPerPixel);   // height
    }

    // ---- Happy path: two-cell overloads -------------------------------------

    [Fact]
    public void AddPicture_TwoCell_Returns_Picture()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        var pic = sh.AddPicture("B2", "D5", OnePixelPng, ImageFormat.Png);
        pic.Format.Should().Be(ImageFormat.Png);
        pic.FromCell.Should().Be("B2");
        pic.ToCell.Should().Be("D5");
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void AddPicture_TwoCell_AutoDetect()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.AddPicture("A1", "C3", OnePixelPng).Format.Should().Be(ImageFormat.Png);
    }

    [Fact]
    public void AddPicture_TwoCell_Uses_A_TwoCellAnchor()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.AddPicture("B2", "D5", OnePixelPng, ImageFormat.Png);

        var dp = wb.Underlying.WorkbookPart!.WorksheetParts
            .SelectMany(p => p.GetPartsOfType<DrawingsPart>()).Single();
        dp.WorksheetDrawing!.Elements<XDR.TwoCellAnchor>().Should().ContainSingle();
    }

    // ---- Validation ---------------------------------------------------------

    [Fact]
    public void AddPicture_Rejects_Null_Cell()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        var act = () => sh.AddPicture(null!, OnePixelPng, ImageFormat.Png);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void AddPicture_Rejects_Null_Data()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        var act = () => sh.AddPicture("A1", null!, ImageFormat.Png);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void AddPicture_Rejects_Invalid_Cell_Address()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        var act = () => sh.AddPicture("notacell", OnePixelPng, ImageFormat.Png);
        act.Should().Throw<InvalidCellAddressException>();
    }

    [Fact]
    public void AddPicture_TwoCell_Rejects_Null_EndCell()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        var act = () => sh.AddPicture("A1", null!, OnePixelPng, ImageFormat.Png);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void AddPicture_AutoDetect_Throws_For_Unknown_Format()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        var gif = new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 }; // 'GIF89a'
        var act = () => sh.AddPicture("A1", gif);
        act.Should().Throw<UnsupportedImageFormatException>()
           .Which.Message.Should().Contain("PNG").And.Contain("JPEG");
    }

    [Fact]
    public void AddPicture_AutoDetect_Throws_For_Garbage_Bytes()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        var act = () => sh.AddPicture("A1", new byte[] { 1, 2, 3 });
        act.Should().Throw<UnsupportedImageFormatException>();
    }

    // ---- Escape hatch (v2.0.0 / I-82) ----------------------------------------

    [Fact]
    public void IPicture_Underlying_Hands_Out_The_Live_Pic_Element()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        var pic = sh.AddPicture("A1", OnePixelPng, ImageFormat.Png);
        // The hatch returns the xdr:pic the add path built — same element the
        // drawing root carries, not a copy.
        pic.Underlying.Should().NotBeNull();
        var root = wb.Underlying.WorkbookPart!.WorksheetParts.Single()
            .GetPartsOfType<DocumentFormat.OpenXml.Packaging.DrawingsPart>().Single().WorksheetDrawing!;
        root.Descendants<XDR.Picture>().Single().Should().BeSameAs(pic.Underlying);
    }

    // ---- Pictures read-back -------------------------------------------------

    [Fact]
    public void Pictures_Empty_When_No_Drawing()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S").Pictures.Should().BeEmpty();
    }

    [Fact]
    public void Pictures_Enumerates_TwoCell_With_Anchor_Offsets_And_Bytes()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddPicture("B3", "D6", OnePixelPng, ImageFormat.Png, dx1: 100, dy1: 200, dx2: 300, dy2: 400);

        var pic = s.Pictures.Should().ContainSingle().Subject;
        pic.Sheet.Should().BeSameAs(s);
        pic.Format.Should().Be(ImageFormat.Png);
        pic.FromCell.Should().Be("B3");
        pic.ToCell.Should().Be("D6");
        pic.Dx1.Should().Be(100); pic.Dy1.Should().Be(200);
        pic.Dx2.Should().Be(300); pic.Dy2.Should().Be(400);
        pic.Data.Should().Equal(OnePixelPng);
    }

    [Fact]
    public void Pictures_Enumerates_OneCell_With_ToCell_Equal_To_FromCell()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddPicture("E5", OnePixelPng, ImageFormat.Png);

        var pic = s.Pictures.Should().ContainSingle().Subject;
        pic.FromCell.Should().Be("E5");
        pic.ToCell.Should().Be("E5");      // one-cell anchor has no distinct end cell
        pic.Dx2.Should().Be(0);
        pic.Dy2.Should().Be(0);
        pic.Data.Should().Equal(OnePixelPng);
    }

    [Fact]
    public void Pictures_Enumerates_Multiple_Mixed_Anchors()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddPicture("A1", OnePixelPng, ImageFormat.Png);          // one-cell
        s.AddPicture("E5", "G7", TinyJpeg, ImageFormat.Jpeg);      // two-cell

        var pics = s.Pictures;
        pics.Should().HaveCount(2);
        pics.Select(p => p.Format).Should().Equal(ImageFormat.Png, ImageFormat.Jpeg);
        pics[0].FromCell.Should().Be("A1");
        pics[1].FromCell.Should().Be("E5");
        pics[1].ToCell.Should().Be("G7");
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- Save / Open round-trip ---------------------------------------------

    [Fact]
    public void Picture_Survives_Save_Open_RoundTrip()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var s = wb.AddSheet("S");
                s.AddPicture("B3", "D6", OnePixelPng, ImageFormat.Png, dx1: 100, dy1: 200, dx2: 300, dy2: 400);
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                var pic = wb["S"].Pictures.Should().ContainSingle().Subject;
                pic.FromCell.Should().Be("B3");
                pic.ToCell.Should().Be("D6");
                pic.Dx1.Should().Be(100); pic.Dy2.Should().Be(400);
                pic.Format.Should().Be(ImageFormat.Png);
                pic.Data.Should().Equal(OnePixelPng);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Multiple_Pictures_Coexist_Across_RoundTrip()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var s = wb.AddSheet("S");
                s.AddPicture("A1", OnePixelPng, ImageFormat.Png);
                s.AddPicture("E5", TinyJpeg, ImageFormat.Jpeg);
                PictureCount(wb, "S").Should().Be(2);
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                wb["S"].Pictures.Should().HaveCount(2);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
