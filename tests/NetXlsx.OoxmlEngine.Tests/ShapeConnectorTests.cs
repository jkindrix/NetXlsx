// I-82 engine swap — drawings slice: shape + connector conformance.
//
// Mirrors the NPOI engine's ISheet.AddShape / AddConnector / Connectors contract on
// the Open XML SDK engine. The NetXlsx.Tests references for parity are ShapeTests,
// ConnectorTests, and the Connectors section of ThemeReadAndDrawingIterationTests.
//
// Per the slice's stated discipline (continuation lesson #6: for geometry, schema-
// valid != positioned-correctly), these tests assert the SDK engine emits the SAME
// ST_ShapeType preset, EMU markers, and <a:ln> line props the NPOI engine emits — not
// just OpenXmlValidator cleanliness. The NPOI oracle XML was captured directly:
//   shape   -> xdr:sp, twoCellAnchor end cell EXCLUSIVE, prstGeom rect/ellipse/...,
//              solidFill|noFill, optional <a:ln><a:solidFill>;
//   connector-> xdr:cxnSp, twoCellAnchor end cell INCLUSIVE, prstGeom
//              straightConnector1/bentConnector3/curvedConnector3, optional <a:ln w=..>
//              with solidFill + head/tail ends, and a fixed <xdr:style> (lnRef idx=1
//              accent1) so LineStyleRefIndex==1 and LineSchemeColor=="accent1".
// The NPOI escape hatch (Underlying) throws NotSupportedException on the SDK engine.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml.Packaging;
using NetXlsx;
using Xunit;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;

namespace NetXlsx.OoxmlEngine.Tests;

public class ShapeConnectorTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-shape-{Guid.NewGuid():N}.xlsx");

    private static XDR.WorksheetDrawing Drawing(IWorkbook wb)
    {
        var dp = wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts
            .SelectMany(p => p.GetPartsOfType<DrawingsPart>()).Single();
        return dp.WorksheetDrawing!;
    }

    // ---- Shapes: happy path -------------------------------------------------

    [Fact]
    public void AddShape_Returns_Shape_With_Type_And_Sheet()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        var shape = s.AddShape(ShapeType.Rectangle, "A1", "D5");
        shape.Should().NotBeNull();
        shape.Type.Should().Be(ShapeType.Rectangle);
        shape.Sheet.Should().BeSameAs(s);
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Theory]
    [InlineData(ShapeType.Rectangle, "rect")]
    [InlineData(ShapeType.RoundedRectangle, "roundRect")]
    [InlineData(ShapeType.Ellipse, "ellipse")]
    [InlineData(ShapeType.Line, "line")]
    [InlineData(ShapeType.Triangle, "triangle")]
    [InlineData(ShapeType.Diamond, "diamond")]
    public void AddShape_Maps_Type_To_NPOI_Preset_Geometry(ShapeType type, string expectedPreset)
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddShape(type, "A1", "C3");

        var sp = Drawing(wb).Descendants<XDR.Shape>().Single();
        sp.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.Preset!.InnerText
            .Should().Be(expectedPreset);
    }

    [Fact]
    public void AddShape_Uses_TwoCellAnchor_With_Exclusive_End_Cell()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddShape(ShapeType.Rectangle, "A1", "D5");

        // NPOI parity: XSSFClientAnchor(0,0,0,0, c1-1, r1-1, c2, r2) — end exclusive.
        var anchor = Drawing(wb).Elements<XDR.TwoCellAnchor>().Single();
        var from = anchor.GetFirstChild<XDR.FromMarker>()!;
        var to = anchor.GetFirstChild<XDR.ToMarker>()!;
        from.ColumnId!.Text.Should().Be("0"); from.RowId!.Text.Should().Be("0");
        to.ColumnId!.Text.Should().Be("4");   // D(3) + 1 exclusive
        to.RowId!.Text.Should().Be("5");       // row 5 (0-based 4) + 1 exclusive
    }

    [Fact]
    public void AddShape_With_FillColor_Emits_SolidFill()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddShape(ShapeType.Ellipse, "B2", "E8", fillColor: Color.FromRgb(255, 0, 0));

        var sp = Drawing(wb).Descendants<XDR.Shape>().Single();
        sp.ShapeProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value
            .Should().Be("FF0000");
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void AddShape_Without_FillColor_Emits_NoFill()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddShape(ShapeType.Rectangle, "A1", "C3");

        var sp = Drawing(wb).Descendants<XDR.Shape>().Single();
        sp.ShapeProperties!.GetFirstChild<A.NoFill>().Should().NotBeNull();
        sp.ShapeProperties!.GetFirstChild<A.SolidFill>().Should().BeNull();
    }

    [Fact]
    public void AddShape_With_LineColor_Emits_Outline_SolidFill()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddShape(ShapeType.Line, "A1", "F1", lineColor: Color.FromRgb(0, 0, 255));

        var sp = Drawing(wb).Descendants<XDR.Shape>().Single();
        sp.ShapeProperties!.GetFirstChild<A.Outline>()!
            .GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value.Should().Be("0000FF");
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void AddShape_Multiple_Coexist_With_Unique_Ids()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddShape(ShapeType.Rectangle, "A1", "C3");
        s.AddShape(ShapeType.Ellipse, "D1", "F3");
        s.AddShape(ShapeType.Triangle, "A4", "C6");

        var shapes = Drawing(wb).Descendants<XDR.Shape>().ToList();
        shapes.Should().HaveCount(3);
        shapes.Select(sp => sp.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id!.Value)
            .Should().OnlyHaveUniqueItems();
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void AddShape_Survives_Save_Open_RoundTrip()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                wb.AddSheet("S").AddShape(ShapeType.Rectangle, "B2", "D5",
                    fillColor: Color.FromRgb(0, 255, 0));
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                Drawing(wb).Descendants<XDR.Shape>().Should().ContainSingle();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void AddShape_Rejects_Null_StartCell()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        var act = () => s.AddShape(ShapeType.Rectangle, null!, "C3");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void AddShape_Rejects_Null_EndCell()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        var act = () => s.AddShape(ShapeType.Rectangle, "A1", null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void IShape_Underlying_Throws_On_The_SDK_Engine()
    {
        using var wb = Workbook.CreateOoxml();
        var shape = wb.AddSheet("S").AddShape(ShapeType.Rectangle, "A1", "C3");
        var act = () => shape.Underlying;
        act.Should().Throw<NotSupportedException>().Which.Message.Should().Contain("OpenXmlDocument");
    }

    // ---- Connectors: happy path ---------------------------------------------

    [Fact]
    public void AddConnector_Returns_Connector_With_Type_And_Sheet()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        var c = s.AddConnector(ConnectorType.Straight, "A1", "C3");
        c.Should().NotBeNull();
        c.Type.Should().Be(ConnectorType.Straight);
        c.Sheet.Should().BeSameAs(s);
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Theory]
    [InlineData(ConnectorType.Straight, "straightConnector1")]
    [InlineData(ConnectorType.Bent, "bentConnector3")]
    [InlineData(ConnectorType.Curved, "curvedConnector3")]
    public void AddConnector_Maps_Type_To_NPOI_Preset_Geometry(ConnectorType type, string expectedPreset)
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConnector(type, "A1", "C3");

        var cxn = Drawing(wb).Descendants<XDR.ConnectionShape>().Single();
        cxn.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.Preset!.InnerText
            .Should().Be(expectedPreset);
    }

    [Fact]
    public void AddConnector_Uses_TwoCellAnchor_With_Inclusive_End_Cell()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        // NPOI parity: connector anchor maps the end cell with -1, so a same-cell
        // connector (I41->I41) has from == to (both at col 8, row 40).
        s.AddConnector(ConnectorType.Straight, "I41", "I41",
            dx1: 115661, dy1: 272142, dx2: 891268, dy2: 278946);

        var anchor = Drawing(wb).Elements<XDR.TwoCellAnchor>().Single();
        var from = anchor.GetFirstChild<XDR.FromMarker>()!;
        var to = anchor.GetFirstChild<XDR.ToMarker>()!;
        from.ColumnId!.Text.Should().Be("8"); from.RowId!.Text.Should().Be("40");
        from.ColumnOffset!.Text.Should().Be("115661"); from.RowOffset!.Text.Should().Be("272142");
        to.ColumnId!.Text.Should().Be("8"); to.RowId!.Text.Should().Be("40");
        to.ColumnOffset!.Text.Should().Be("891268"); to.RowOffset!.Text.Should().Be("278946");
    }

    [Fact]
    public void AddConnector_Emits_Line_Width_Flip_And_TailEnd()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConnector(ConnectorType.Straight, "I41", "I41",
            lineColor: Color.FromRgb(0, 0, 0),
            flipH: true, tailEnd: ConnectorEnd.Arrow, lineWidthPoints: 2.0);

        var spPr = Drawing(wb).Descendants<XDR.ConnectionShape>().Single().ShapeProperties!;
        spPr.Transform2D!.HorizontalFlip!.Value.Should().BeTrue();
        var ln = spPr.GetFirstChild<A.Outline>()!;
        ln.Width!.Value.Should().Be(25400);   // 2.0pt * 12700 EMU/pt
        ln.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value.Should().Be("000000");
        ln.GetFirstChild<A.TailEnd>()!.Type!.InnerText.Should().Be("arrow");
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void AddConnector_Emits_The_NPOI_Style_Block()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConnector(ConnectorType.Straight, "A1", "C3");

        var style = Drawing(wb).Descendants<XDR.ConnectionShape>().Single().ShapeStyle!;
        style.LineReference!.Index!.Value.Should().Be(1U);
        style.LineReference!.SchemeColor!.Val!.InnerText.Should().Be("accent1");
    }

    [Fact]
    public void AddConnector_Survives_Save_Open_RoundTrip()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                wb.AddSheet("S").AddConnector(ConnectorType.Straight, "I41", "I41",
                    lineColor: Color.FromRgb(0, 0, 0),
                    dx1: 115661, dy1: 272142, dx2: 891268, dy2: 278946,
                    flipH: true, tailEnd: ConnectorEnd.Arrow, lineWidthPoints: 2.0);
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                var c = wb["S"].Connectors.Should().ContainSingle().Subject;
                c.Type.Should().Be(ConnectorType.Straight);
                c.FromCell.Should().Be("I41");
                c.ToCell.Should().Be("I41");
                c.Dx1.Should().Be(115661); c.Dy2.Should().Be(278946);
                c.FlipH.Should().BeTrue();
                c.TailEnd.Should().Be(ConnectorEnd.Arrow);
                c.LineColor.Should().Be(Color.FromRgb(0, 0, 0));
                c.LineWidthPoints.Should().BeApproximately(2.0, 0.001);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void AddConnector_Rejects_Null_StartCell()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        var act = () => s.AddConnector(ConnectorType.Straight, null!, "C3");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void IConnector_Underlying_Throws_On_The_SDK_Engine()
    {
        using var wb = Workbook.CreateOoxml();
        var c = wb.AddSheet("S").AddConnector(ConnectorType.Straight, "A1", "C3");
        var act = () => c.Underlying;
        act.Should().Throw<NotSupportedException>().Which.Message.Should().Contain("OpenXmlDocument");
    }

    // ---- Connectors read-back -----------------------------------------------

    [Fact]
    public void Connectors_Empty_When_No_Drawing()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S").Connectors.Should().BeEmpty();
    }

    [Fact]
    public void Connectors_Enumerates_Added_Connectors_With_All_Properties()
    {
        // Mirrors ThemeReadAndDrawingIterationTests on the NPOI engine.
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConnector(ConnectorType.Straight, "I41", "I41",
            lineColor: Color.FromRgb(0, 0, 0),
            dx1: 115661, dy1: 272142, dx2: 891268, dy2: 278946,
            flipH: true, tailEnd: ConnectorEnd.Arrow, lineWidthPoints: 2.0);

        var c = s.Connectors.Should().ContainSingle().Subject;
        c.Type.Should().Be(ConnectorType.Straight);
        c.FromCell.Should().Be("I41");
        c.ToCell.Should().Be("I41");
        c.Dx1.Should().Be(115661); c.Dy2.Should().Be(278946);
        c.FlipH.Should().BeTrue();
        c.FlipV.Should().BeFalse();
        c.HeadEnd.Should().Be(ConnectorEnd.None);
        c.TailEnd.Should().Be(ConnectorEnd.Arrow);
        c.LineColor.Should().Be(Color.FromRgb(0, 0, 0));
        c.LineWidthPoints.Should().BeApproximately(2.0, 0.001);
        c.LineSchemeColor.Should().Be("accent1");   // from the style block's lnRef
        c.LineStyleRefIndex.Should().Be(1);
    }

    [Fact]
    public void Connectors_Enumerates_Multiple_In_Order()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddConnector(ConnectorType.Straight, "A1", "C3");
        s.AddConnector(ConnectorType.Bent, "D1", "F3");
        s.AddConnector(ConnectorType.Curved, "G1", "I3");

        s.Connectors.Select(c => c.Type).Should().Equal(
            ConnectorType.Straight, ConnectorType.Bent, ConnectorType.Curved);
    }

    // ---- Cross-contamination guards -----------------------------------------
    //
    // Pictures / Connectors each filter the shared xdr:wsDr for their own anchor
    // child kind; shapes have no read-back. These guard that the three coexist
    // without one enumeration picking up another's anchors.

    [Fact]
    public void Connectors_Ignores_Shapes_And_Pictures()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddShape(ShapeType.Rectangle, "A1", "C3");
        s.AddPicture("E5", PictureBytes.OnePixelPng, ImageFormat.Png);
        s.AddConnector(ConnectorType.Straight, "A8", "C10");

        s.Connectors.Should().ContainSingle().Which.Type.Should().Be(ConnectorType.Straight);
    }

    [Fact]
    public void Pictures_Ignores_Shapes_And_Connectors()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.AddShape(ShapeType.Rectangle, "A1", "C3");
        s.AddConnector(ConnectorType.Straight, "A8", "C10");
        s.AddPicture("E5", PictureBytes.OnePixelPng, ImageFormat.Png);

        s.Pictures.Should().ContainSingle().Which.FromCell.Should().Be("E5");
        OpenXmlValidationGate.AssertValid(wb);
    }

    private static class PictureBytes
    {
        public static readonly byte[] OnePixelPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");
    }
}
