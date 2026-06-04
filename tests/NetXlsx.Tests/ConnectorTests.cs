using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class ConnectorTests
{
    [Fact]
    public void AddConnector_Straight_Uses_Correct_Preset_Geometry()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var c = s.AddConnector(ConnectorType.Straight, "A1", "C3");

        c.Should().NotBeNull();
        c.Type.Should().Be(ConnectorType.Straight);
        c.Sheet.Should().BeSameAs(s);
        PresetGeometryOf(wb).Should().Be("straightConnector1");
    }

    [Theory]
    [InlineData(ConnectorType.Straight, "straightConnector1")]
    [InlineData(ConnectorType.Bent, "bentConnector3")]
    [InlineData(ConnectorType.Curved, "curvedConnector3")]
    public void AddConnector_Maps_Type_To_Geometry(ConnectorType type, string expectedPreset)
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConnector(type, "A1", "C3");
        PresetGeometryOf(wb).Should().Be(expectedPreset);
    }

    [Fact]
    public void AddConnector_Sets_Offsets_FlipH_And_TailArrow()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var c = s.AddConnector(ConnectorType.Straight, "I41", "I41",
            lineColor: Color.FromRgb(0, 0, 0),
            dx1: 115661, dy1: 272142, dx2: 891268, dy2: 278946,
            flipH: true, tailEnd: ConnectorEnd.Arrow,
            lineWidthPoints: 2.0);

        var anchor = SavedOoxml.DrawingXml(wb).Root!
            .Element(SavedOoxml.Xdr + "twoCellAnchor")!;
        var from = anchor.Element(SavedOoxml.Xdr + "from")!;
        ((int)from.Element(SavedOoxml.Xdr + "col")!).Should().Be(8);
        ((int)from.Element(SavedOoxml.Xdr + "row")!).Should().Be(40);
        ((long)from.Element(SavedOoxml.Xdr + "colOff")!).Should().Be(115661);
        ((long)anchor.Element(SavedOoxml.Xdr + "to")!
            .Element(SavedOoxml.Xdr + "colOff")!).Should().Be(891268);

        var spPr = anchor.Descendants(SavedOoxml.Xdr + "spPr").Single();
        SavedOoxml.BoolAttr(spPr.Element(SavedOoxml.Dml + "xfrm"), "flipH").Should().BeTrue();
        ((string?)spPr.Element(SavedOoxml.Dml + "ln")!
            .Element(SavedOoxml.Dml + "tailEnd")!.Attribute("type")).Should().Be("arrow");
    }

    [Fact]
    public void AddConnector_Survives_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("S");
            s.AddConnector(ConnectorType.Straight, "I41", "I41",
                lineColor: Color.FromRgb(0, 0, 0),
                dx1: 115661, dy1: 272142, dx2: 891268, dy2: 278946,
                flipH: true, tailEnd: ConnectorEnd.Arrow, lineWidthPoints: 2.0);
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        SavedOoxml.DrawingXml(opened).Descendants(SavedOoxml.Xdr + "cxnSp")
            .Should().HaveCount(1);
    }

    [Fact]
    public void AddConnector_Rejects_Null_StartCell()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.AddConnector(ConnectorType.Straight, null!, "C3");
        act.Should().Throw<ArgumentNullException>();
    }

    // ---- helpers ------------------------------------------------------

    /// <summary>The preset geometry (a:prstGeom/@prst) of the sole connector.</summary>
    private static string? PresetGeometryOf(IWorkbook wb)
        => (string?)SavedOoxml.DrawingXml(wb)
            .Descendants(SavedOoxml.Xdr + "cxnSp").Single()
            .Descendants(SavedOoxml.Dml + "prstGeom").Single()
            .Attribute("prst");
}
