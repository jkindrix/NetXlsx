using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class ConditionalFormatTests
{
    [Fact]
    public void CellValueGreaterThan_Adds_Rule()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.CellValueGreaterThan("50", new CellStyle { Bold = true }));

        s.ConditionalFormattingCount.Should().Be(1);
    }

    [Fact]
    public void CellValueBetween_Adds_Rule()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.CellValueBetween("10", "90", new CellStyle { Italic = true }));

        s.ConditionalFormattingCount.Should().Be(1);
    }

    [Fact]
    public void CellValueEqual_Adds_Rule()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A5",
            ConditionalFormat.CellValueEqual("100", new CellStyle { Bold = true }));

        s.ConditionalFormattingCount.Should().Be(1);
    }

    [Fact]
    public void Formula_Rule_Adds_Successfully()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.Formula("$A1>50", new CellStyle { Bold = true }));

        s.ConditionalFormattingCount.Should().Be(1);
    }

    [Fact]
    public void ColorScale_Two_Color_Adds_Rule()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.ColorScale(Color.FromRgb(255, 0, 0), Color.FromRgb(0, 255, 0)));

        s.ConditionalFormattingCount.Should().Be(1);
    }

    [Fact]
    public void ColorScale_Three_Color_Adds_Rule()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.ColorScale(
                Color.FromRgb(255, 0, 0),
                Color.FromRgb(255, 255, 0),
                Color.FromRgb(0, 255, 0)));

        s.ConditionalFormattingCount.Should().Be(1);
    }

    [Fact]
    public void Multiple_Rules_On_Same_Range()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A10",
            ConditionalFormat.CellValueGreaterThan("80", new CellStyle { Bold = true }),
            ConditionalFormat.CellValueLessThan("20", new CellStyle { Italic = true }));

        s.ConditionalFormattingCount.Should().Be(1);
    }

    [Fact]
    public void Multiple_AddConditionalFormatting_Calls()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A5",
            ConditionalFormat.CellValueGreaterThan("50", new CellStyle { Bold = true }));
        s.AddConditionalFormatting("B1:B5",
            ConditionalFormat.CellValueLessThan("10", new CellStyle { Italic = true }));

        s.ConditionalFormattingCount.Should().Be(2);
    }

    [Fact]
    public void RemoveConditionalFormatting_Decrements_Count()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A5",
            ConditionalFormat.CellValueGreaterThan("50", new CellStyle { Bold = true }));
        s.AddConditionalFormatting("B1:B5",
            ConditionalFormat.CellValueLessThan("10", new CellStyle { Italic = true }));

        s.RemoveConditionalFormatting(0);
        s.ConditionalFormattingCount.Should().Be(1);
    }

    [Fact]
    public void Rejects_Null_Range()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.AddConditionalFormatting(null!,
            ConditionalFormat.CellValueGreaterThan("1", new CellStyle()));
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Rejects_Empty_Rules()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s.AddConditionalFormatting("A1:A5");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void CellValue_Rule_Survives_RoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("S");
            for (int i = 1; i <= 5; i++) s[$"A{i}"].SetNumber(i * 20);
            s.AddConditionalFormatting("A1:A5",
                ConditionalFormat.CellValueGreaterThan("50", new CellStyle { Bold = true }));
            wb.Save(ms, leaveOpen: true);
        }

        ms.Position = 0;
        using var opened = Workbook.Open(ms);
        opened["S"].ConditionalFormattingCount.Should().Be(1);
    }

    [Fact]
    public void CellValueNotEqual_Adds_Rule()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A5",
            ConditionalFormat.CellValueNotEqual("0", new CellStyle { Bold = true }));
        s.ConditionalFormattingCount.Should().Be(1);
    }

    [Fact]
    public void CellValueGreaterThanOrEqual_Adds_Rule()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A5",
            ConditionalFormat.CellValueGreaterThanOrEqual("50", new CellStyle { Bold = true }));
        s.ConditionalFormattingCount.Should().Be(1);
    }

    [Fact]
    public void CellValueLessThanOrEqual_Adds_Rule()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.AddConditionalFormatting("A1:A5",
            ConditionalFormat.CellValueLessThanOrEqual("50", new CellStyle { Bold = true }));
        s.ConditionalFormattingCount.Should().Be(1);
    }
}
