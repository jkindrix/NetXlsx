using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class ShapeTests
{
    [Fact]
    public void AddShape_Rectangle_Creates_Shape()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var shape = s.AddShape(ShapeType.Rectangle, "A1", "D5");

        shape.Should().NotBeNull();
        shape.Type.Should().Be(ShapeType.Rectangle);
        shape.Sheet.Should().BeSameAs(s);
        shape.Underlying.Should().NotBeNull();
    }

    [Fact]
    public void AddShape_With_FillColor()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var shape = s.AddShape(ShapeType.Ellipse, "B2", "E8",
            fillColor: Color.FromRgb(255, 0, 0));

        shape.Type.Should().Be(ShapeType.Ellipse);
    }

    [Fact]
    public void AddShape_With_LineColor()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var shape = s.AddShape(ShapeType.Line, "A1", "F1",
            lineColor: Color.FromRgb(0, 0, 255));

        shape.Type.Should().Be(ShapeType.Line);
    }

    [Fact]
    public void AddShape_Multiple_Shapes_On_Same_Sheet()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddShape(ShapeType.Rectangle, "A1", "C3");
        s.AddShape(ShapeType.Ellipse, "D1", "F3");
        s.AddShape(ShapeType.Triangle, "A4", "C6");

        SavedOoxml.DrawingXml(wb).Descendants(SavedOoxml.Xdr + "sp")
            .Should().HaveCount(3);
    }

    [Fact]
    public void AddShape_RoundedRectangle()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var shape = s.AddShape(ShapeType.RoundedRectangle, "A1", "D4",
            fillColor: Color.FromRgb(0, 128, 255));

        shape.Type.Should().Be(ShapeType.RoundedRectangle);
    }

    [Fact]
    public void AddShape_Survives_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("S");
            s.AddShape(ShapeType.Rectangle, "B2", "D5",
                fillColor: Color.FromRgb(0, 255, 0));
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        SavedOoxml.DrawingXml(opened).Descendants(SavedOoxml.Xdr + "sp")
            .Should().HaveCount(1);
    }

    [Fact]
    public void AddShape_Rejects_Null_StartCell()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.AddShape(ShapeType.Rectangle, null!, "C3");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void AddShape_Rejects_Null_EndCell()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.AddShape(ShapeType.Rectangle, "A1", null!);
        act.Should().Throw<ArgumentNullException>();
    }
}
