using System;
using System.IO;
using AwesomeAssertions;
using NPOI.OpenXmlFormats.Dml;
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
        c.Underlying.GetCTConnector().spPr.prstGeom.prst
            .Should().Be(ST_ShapeType.straightConnector1);
    }

    [Theory]
    [InlineData(ConnectorType.Straight, ST_ShapeType.straightConnector1)]
    [InlineData(ConnectorType.Bent, ST_ShapeType.bentConnector3)]
    [InlineData(ConnectorType.Curved, ST_ShapeType.curvedConnector3)]
    public void AddConnector_Maps_Type_To_Geometry(ConnectorType type, ST_ShapeType expected)
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var c = s.AddConnector(type, "A1", "C3");
        c.Underlying.GetCTConnector().spPr.prstGeom.prst.Should().Be(expected);
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

        var anchor = (NPOI.XSSF.UserModel.XSSFClientAnchor)c.Underlying.GetAnchor();
        anchor.Dx1.Should().Be(115661);
        anchor.Dx2.Should().Be(891268);
        anchor.Col1.Should().Be(8);
        anchor.Row1.Should().Be(40);

        var ct = c.Underlying.GetCTConnector();
        ct.spPr.xfrm.flipH.Should().BeTrue();
        ct.spPr.ln.tailEnd.type.Should().Be(ST_LineEndType.arrow);
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
        var drawing = (NPOI.XSSF.UserModel.XSSFDrawing)opened["S"].Underlying.CreateDrawingPatriarch();
        drawing.GetShapes().Count.Should().Be(1);
    }

    [Fact]
    public void AddConnector_Rejects_Null_StartCell()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.AddConnector(ConnectorType.Straight, null!, "C3");
        act.Should().Throw<ArgumentNullException>();
    }
}
