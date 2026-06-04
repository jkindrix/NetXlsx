// Coverage for the v1.1 image-embedding slice: ISheet.AddPicture
// (explicit format + auto-detect overloads), magic-byte detection,
// round-trip preservation through Workbook.Open.

using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class PictureApiTests
{
    // Known-valid 1×1 transparent PNG (the canonical "tiny PNG" — 67
    // bytes encoded base64). Loaded once at class init.
    private static readonly byte[] s_tinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk" +
        "YAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

    // Minimal JPEG: SOI + DQT + SOF0 + DHT + SOS + entropy + EOI.
    // This is the smallest valid JPEG that Excel will load (a 1×1
    // grayscale JPEG ≈ 125 bytes). The leading FF D8 FF is what the
    // magic-byte detector checks.
    private static readonly byte[] s_tinyJpeg = new byte[]
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

    // ---- Happy path ----------------------------------------------------

    [Fact]
    public void AddPicture_With_Explicit_Format_Returns_Picture()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        var pic = sh.AddPicture("B2", s_tinyPng, ImageFormat.Png);
        pic.Sheet.Should().BeSameAs(sh);
        pic.Format.Should().Be(ImageFormat.Png);
        pic.Underlying.Should().NotBeNull();
    }

    [Fact]
    public void AddPicture_AutoDetects_PNG_From_Magic_Bytes()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        var pic = sh.AddPicture("A1", s_tinyPng);
        pic.Format.Should().Be(ImageFormat.Png);
    }

    [Fact]
    public void AddPicture_AutoDetects_JPEG_From_Magic_Bytes()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        var pic = sh.AddPicture("A1", s_tinyJpeg);
        pic.Format.Should().Be(ImageFormat.Jpeg);
    }

    // ---- Validation ----------------------------------------------------

    [Fact]
    public void AddPicture_Rejects_Null_Cell()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.AddPicture(null!, s_tinyPng, ImageFormat.Png);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void AddPicture_Rejects_Null_Data()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.AddPicture("A1", null!, ImageFormat.Png);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void AddPicture_Rejects_Invalid_Cell_Address()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.AddPicture("notacell", s_tinyPng, ImageFormat.Png);
        act.Should().Throw<InvalidCellAddressException>();
    }

    [Fact]
    public void AddPicture_AutoDetect_Throws_For_Unknown_Format()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        // GIF magic 'GIF8' — not in our v1.1 set.
        var gif = new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 };
        Action act = () => sh.AddPicture("A1", gif);
        act.Should().Throw<UnsupportedImageFormatException>()
           .Which.Message.Should().Contain("PNG").And.Contain("JPEG");
    }

    [Fact]
    public void AddPicture_AutoDetect_Throws_For_Garbage_Bytes()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.AddPicture("A1", new byte[] { 1, 2, 3 });
        act.Should().Throw<UnsupportedImageFormatException>();
    }

    // ---- File round-trip ----------------------------------------------

    [Fact]
    public void AddPicture_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"pic-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh.AddPicture("B2", s_tinyPng, ImageFormat.Png);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var sh = wb["S"];
                sh.Pictures.Should().HaveCount(1);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void AddPicture_Multiple_Pictures_Coexist()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.AddPicture("A1", s_tinyPng, ImageFormat.Png);
        sh.AddPicture("E5", s_tinyJpeg, ImageFormat.Jpeg);
        sh.Pictures.Should().HaveCount(2);
    }
}
