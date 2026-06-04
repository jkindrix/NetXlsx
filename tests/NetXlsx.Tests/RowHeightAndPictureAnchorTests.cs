using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class RowHeightAndPictureAnchorTests
{
    [Fact]
    public void HeightInPoints_Set_And_Get()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var row = s.Row(1);
        row.HeightInPoints = 24.95f;
        row.HeightInPoints.Should().BeApproximately(24.95f, 0.1f);
    }

    [Fact]
    public void HeightInPoints_Default_Is_Sheet_Default()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var row = s.Row(1);
        row.HeightInPoints.Should().BeApproximately(15f, 1f);
    }

    [Fact]
    public void HeightInPoints_Survives_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("S");
            s.Row(1).HeightInPoints = 30f;
            s.Row(2).HeightInPoints = 5.25f;
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        opened["S"].Row(1).HeightInPoints.Should().BeApproximately(30f, 0.1f);
        opened["S"].Row(2).HeightInPoints.Should().BeApproximately(5.25f, 0.1f);
    }

    private static byte[] MakePng()
    {
        // Minimal valid 1x1 PNG
        return new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
            0x44, 0xAE, 0x42, 0x60, 0x82
        };
    }

    [Fact]
    public void AddPicture_TwoCell_Anchor()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var pic = s.AddPicture("B2", "D5", MakePng(), ImageFormat.Png);
        pic.Should().NotBeNull();
        pic.Format.Should().Be(ImageFormat.Png);
    }

    [Fact]
    public void AddPicture_TwoCell_AutoDetect()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var pic = s.AddPicture("A1", "C3", MakePng());
        pic.Format.Should().Be(ImageFormat.Png);
    }

    [Fact]
    public void AddPicture_TwoCell_Survives_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("S");
            s.AddPicture("B7", "C9", MakePng(), ImageFormat.Png);
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        SavedOoxml.DrawingXml(opened).Descendants(SavedOoxml.Xdr + "pic")
            .Should().HaveCount(1);
    }

    [Fact]
    public void AddPicture_TwoCell_Rejects_Null_EndCell()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.AddPicture("A1", null!, MakePng(), ImageFormat.Png);
        act.Should().Throw<ArgumentNullException>();
    }
}
