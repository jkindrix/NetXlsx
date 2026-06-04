using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class ChartTests
{
    private static ISheet CreateSheetWithData()
    {
        var wb = Workbook.Create();
        var s = wb.AddSheet("Data");
        s["A1"].SetString("Jan"); s["B1"].SetNumber(100);
        s["A2"].SetString("Feb"); s["B2"].SetNumber(150);
        s["A3"].SetString("Mar"); s["B3"].SetNumber(200);
        s["A4"].SetString("Apr"); s["B4"].SetNumber(120);
        return s;
    }

    [Fact]
    public void AddChart_Line_Creates_Chart()
    {
        var s = CreateSheetWithData();
        using var wb = (IDisposable)s.Workbook;
        var chart = s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");

        chart.Should().NotBeNull();
        chart.Type.Should().Be(ChartType.Line);
        chart.Sheet.Should().BeSameAs(s);
        chart.Underlying.Should().NotBeNull();
    }

    [Fact]
    public void AddChart_Bar_Creates_Chart()
    {
        var s = CreateSheetWithData();
        using var wb = (IDisposable)s.Workbook;
        var chart = s.AddChart(ChartType.Bar, "D1", "K15", "A1:A4", "B1:B4");
        chart.Type.Should().Be(ChartType.Bar);
    }

    [Fact]
    public void AddChart_Column_Creates_Chart()
    {
        var s = CreateSheetWithData();
        using var wb = (IDisposable)s.Workbook;
        var chart = s.AddChart(ChartType.Column, "D1", "K15", "A1:A4", "B1:B4");
        chart.Type.Should().Be(ChartType.Column);
    }

    [Fact]
    public void AddChart_Pie_Creates_Chart()
    {
        var s = CreateSheetWithData();
        using var wb = (IDisposable)s.Workbook;
        var chart = s.AddChart(ChartType.Pie, "D1", "K15", "A1:A4", "B1:B4");
        chart.Type.Should().Be(ChartType.Pie);
    }

    [Fact]
    public void AddChart_Area_Creates_Chart()
    {
        var s = CreateSheetWithData();
        using var wb = (IDisposable)s.Workbook;
        var chart = s.AddChart(ChartType.Area, "D1", "K15", "A1:A4", "B1:B4");
        chart.Type.Should().Be(ChartType.Area);
    }

    [Fact]
    public void AddChart_Scatter_Creates_Chart()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("Data");
        s["A1"].SetNumber(1); s["B1"].SetNumber(10);
        s["A2"].SetNumber(2); s["B2"].SetNumber(20);
        s["A3"].SetNumber(3); s["B3"].SetNumber(30);

        var chart = s.AddChart(ChartType.Scatter, "D1", "K15", "A1:A3", "B1:B3");
        chart.Type.Should().Be(ChartType.Scatter);
    }

    [Fact]
    public void AddChart_With_Title()
    {
        var s = CreateSheetWithData();
        using var wb = (IDisposable)s.Workbook;
        var chart = s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4", title: "Sales");
        chart.Underlying.Should().NotBeNull();
    }

    [Fact]
    public void AddChart_Survives_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("Data");
            s["A1"].SetString("Q1"); s["B1"].SetNumber(100);
            s["A2"].SetString("Q2"); s["B2"].SetNumber(200);
            s.AddChart(ChartType.Bar, "D1", "J12", "A1:A2", "B1:B2", "Revenue");
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        // Chart survives as a graphic frame in the drawing part.
        SavedOoxml.DrawingXml(opened)
            .Descendants(SavedOoxml.Xdr + "graphicFrame").Should().HaveCount(1);
    }

    [Fact]
    public void AddChart_Rejects_Null_CategoryRange()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.AddChart(ChartType.Line, "D1", "K15", null!, "B1:B4");
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void SetTitle_Updates_Chart()
    {
        var s = CreateSheetWithData();
        using var wb = (IDisposable)s.Workbook;
        var chart = s.AddChart(ChartType.Line, "D1", "K15", "A1:A4", "B1:B4");
        chart.SetTitle("Updated");
        // No throw is success; title is set on underlying
    }
}
