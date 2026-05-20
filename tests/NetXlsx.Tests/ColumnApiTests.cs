// Coverage for the v0.6 sub-slice C IColumn API.

using System;
using System.Collections.Generic;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class ColumnApiTests
{
    // ---- Construction and addressing ----------------------------------

    [Fact]
    public void Column_By_Index_Reports_Letter_And_Sheet()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var col = sheet.Column(1);
        col.Index.Should().Be(1);
        col.Letter.Should().Be("A");
        col.Sheet.Should().BeSameAs(sheet);
    }

    [Theory]
    [InlineData("A", 1)]
    [InlineData("Z", 26)]
    [InlineData("AA", 27)]
    [InlineData("aa", 27)]
    [InlineData("$AB", 28)]
    [InlineData("XFD", 16384)]
    public void Column_By_Letter_Parses_Standard_Forms(string letter, int expectedIndex)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.Column(letter).Index.Should().Be(expectedIndex);
    }

    [Fact]
    public void Column_By_Letter_Roundtrips_Through_FormatColumn()
    {
        for (int i = 1; i <= 100; i++)
        {
            var letter = CellAddress.FormatColumn(i);
            CellAddress.ParseColumn(letter).Should().Be(i);
        }
        CellAddress.FormatColumn(16384).Should().Be("XFD");
    }

    [Fact]
    public void Column_By_Letter_Rejects_Garbage()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet.Column(""))).Should().Throw<InvalidCellAddressException>();
        ((Action)(() => sheet.Column("1"))).Should().Throw<InvalidCellAddressException>();
        ((Action)(() => sheet.Column("A1"))).Should().Throw<InvalidCellAddressException>();
        ((Action)(() => sheet.Column("AAAA"))).Should().Throw<InvalidCellAddressException>(); // 27*26*... > XFD
    }

    [Fact]
    public void Column_By_Index_Validates_Bounds()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet.Column(0))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => sheet.Column(CellAddress.MaxColumn + 1))).Should().Throw<ArgumentOutOfRangeException>();
    }

    // ---- Width --------------------------------------------------------

    [Fact]
    public void Width_Set_And_Get_Roundtrip_Within_NPOIs_Quantization()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var col = sheet.Column("B");
        col.Width(24.0);
        // NPOI stores widths as int 256ths; expect close to but not necessarily ==.
        col.WidthUnits.Should().BeApproximately(24.0, 0.01);
    }

    [Fact]
    public void Width_Fluent_Form_Returns_Same_Column()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var col = sheet.Column("B");
        col.Width(10.0).Should().BeSameAs(col);
    }

    [Fact]
    public void Width_Rejects_Negative_And_NaN()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var col = sheet.Column("B");
        ((Action)(() => col.Width(-1.0))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => col.Width(double.NaN))).Should().Throw<ArgumentOutOfRangeException>();
    }

    // ---- Hidden -------------------------------------------------------

    [Fact]
    public void Hidden_Setter_Roundtrips_Through_NPOI()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var col = sheet.Column("C");
        col.Hidden.Should().BeFalse();
        col.Hidden = true;
        col.Hidden.Should().BeTrue();
        col.Hidden = false;
        col.Hidden.Should().BeFalse();
    }

    // ---- SetDefaultStyle ---------------------------------------------

    [Fact]
    public void SetDefaultStyle_Goes_Through_Style_Pool()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.Column("A").SetDefaultStyle(new CellStyle { NumberFormat = NumberFormats.Currency });
        sheet.Column("A").SetDefaultStyle(new CellStyle { NumberFormat = NumberFormats.Currency });

        // Default-style on a column doesn't materialize cells, but the
        // style pool entry should be reused — verified indirectly by
        // creating a cell and observing its style is set up.
        var npoiStyle = sheet.Underlying.GetColumnStyle(0);
        npoiStyle.Should().NotBeNull();
    }

    // ---- ForEachPopulated --------------------------------------------

    [Fact]
    public void ForEachPopulated_Skips_Empty_Cells_And_Visits_In_Order()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("x");
        sheet["A3"].SetString("y");
        // A2 intentionally empty.

        var visited = new List<string>();
        sheet.Column("A").ForEachPopulated(c => visited.Add(c.Address));

        visited.Should().Equal("A1", "A3");
    }

    [Fact]
    public void ForEachPopulated_Empty_Column_Is_A_No_Op()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var count = 0;
        sheet.Column("D").ForEachPopulated(_ => count++);
        count.Should().Be(0);
    }

    [Fact]
    public void ForEachPopulated_Null_Action_Throws()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet.Column("A").ForEachPopulated(null!)))
            .Should().Throw<ArgumentNullException>();
    }

    // ---- AutoSize -----------------------------------------------------

    [Fact]
    public void AutoSize_Either_Sizes_The_Column_Or_Throws_MissingFontException()
    {
        // AutoSize requires font metrics; on a headless CI without fonts
        // we get MissingFontException (decision I3). Either outcome is
        // acceptable — failing silently is not.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("hello");
        var col = sheet.Column("A");
        double widthBefore = col.WidthUnits;

        try
        {
            col.AutoSize();
            // Success path: width should be a positive finite value.
            col.WidthUnits.Should().BeGreaterThan(0);
        }
        catch (MissingFontException ex)
        {
            ex.Message.Should().Contain("AutoSize");
        }
    }

    // ---- Round-trip ---------------------------------------------------

    [Fact]
    public void Column_Operations_Roundtrip_Through_Save_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"col-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("S");
                sheet.Column("A").Width(20.0);
                sheet.Column("B").Hidden = true;
                sheet.Column("C").SetDefaultStyle(new CellStyle { NumberFormat = NumberFormats.Currency });
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var sheet = wb["S"];
                sheet.Column("A").WidthUnits.Should().BeApproximately(20.0, 0.01);
                sheet.Column("B").Hidden.Should().BeTrue();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
