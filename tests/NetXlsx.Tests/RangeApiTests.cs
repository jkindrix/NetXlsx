// Coverage for the v0.6 sub-slice B IRange API.

using System;
using System.IO;
using System.Linq;
using FluentAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class RangeApiTests
{
    private static readonly string[] SparseExpected = new[] { "A1", "C3" };
    private static readonly string[] DenseExpected  = new[] { "A1", "A2", "B1", "B2" };

    // ---- Construction and addressing -----------------------------------

    [Fact]
    public void Range_A1Range_Roundtrips_Through_Address()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var r = sheet.Range("A1:C3");
        r.Address.Should().Be("A1:C3");
        r.FirstRow.Should().Be(1);
        r.FirstCol.Should().Be(1);
        r.LastRow.Should().Be(3);
        r.LastCol.Should().Be(3);
        r.Count.Should().Be(9, "dense coordinate count = 3 * 3");
    }

    [Fact]
    public void Range_Coordinate_Form_Normalizes_Inverted_Corners()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var r = sheet.Range(3, 3, 1, 1);
        r.Address.Should().Be("A1:C3");
        r.FirstRow.Should().Be(1);
    }

    [Fact]
    public void Range_Coordinate_Form_Validates_Bounds()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet.Range(0, 1, 1, 1))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => sheet.Range(1, 0, 1, 1))).Should().Throw<ArgumentOutOfRangeException>();
        ((Action)(() => sheet.Range(1, 1, CellAddress.MaxRow + 1, 1))).Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Range_Single_Cell_Form_Has_Count_One()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var r = sheet.Range("B2");
        r.Address.Should().Be("B2");
        r.Count.Should().Be(1);
    }

    [Fact]
    public void Range_WholeColumn_Form_Expands_To_MaxRow()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var r = sheet.Range("A:A");
        r.FirstRow.Should().Be(1);
        r.LastRow.Should().Be(CellAddress.MaxRow);
        r.FirstCol.Should().Be(1);
        r.LastCol.Should().Be(1);
    }

    [Fact]
    public void Range_WholeRow_Form_Expands_To_MaxColumn()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var r = sheet.Range("3:3");
        r.FirstRow.Should().Be(3);
        r.LastRow.Should().Be(3);
        r.FirstCol.Should().Be(1);
        r.LastCol.Should().Be(CellAddress.MaxColumn);
    }

    // ---- Value(object?) ------------------------------------------------

    [Theory]
    [InlineData("text",          CellKind.String)]
    [InlineData((int)42,         CellKind.Number)]
    [InlineData((long)42L,       CellKind.Number)]
    [InlineData(2.5,             CellKind.Number)]
    [InlineData(true,            CellKind.Bool)]
    public void Value_Dispatches_On_Runtime_Type_And_Fills_All_Cells(object value, CellKind expectedKind)
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.Range("A1:C2").Value(value);

        for (int r = 1; r <= 2; r++)
            for (int c = 1; c <= 3; c++)
                sheet[r, c].Kind.Should().Be(expectedKind, $"cell at ({r},{c}) should have kind {expectedKind}");
    }

    [Fact]
    public void Value_Decimal_Fills_As_Number()
    {
        // decimal can't be used as an InlineData parameter — separate test.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.Range("A1:B1").Value(1234.56m);
        sheet["A1"].GetNumber().Should().Be(1234.56);
        sheet["B1"].GetNumber().Should().Be(1234.56);
    }

    [Fact]
    public void Value_DateTime_Fills_As_Date()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var dt = new DateTime(2026, 5, 16, 0, 0, 0, DateTimeKind.Unspecified);
        sheet.Range("A1:B1").Value(dt);
        sheet["A1"].GetDate().Should().Be(dt);
        sheet["B1"].Kind.Should().Be(CellKind.Date);
    }

    [Fact]
    public void Value_Null_Clears_Every_Cell()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.Range("A1:C2").Value("filler");
        sheet.Range("A1:C2").Value(null);
        for (int r = 1; r <= 2; r++)
            for (int c = 1; c <= 3; c++)
                sheet[r, c].Kind.Should().Be(CellKind.Empty);
    }

    [Fact]
    public void Value_Unsupported_Type_Throws_ArgumentException()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => sheet.Range("A1").Value(new System.Text.StringBuilder())))
            .Should().Throw<ArgumentException>()
            .WithMessage("*StringBuilder*not a supported scalar*");
    }

    // ---- Apply(CellStyle) ---------------------------------------------

    [Fact]
    public void Apply_Sets_Style_On_Every_Cell_In_Rectangle()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.Range("A1:B2").Value(1);
        sheet.Range("A1:B2").Apply(new CellStyle { Bold = true, Background = Color.Yellow });

        for (int r = 1; r <= 2; r++)
            for (int c = 1; c <= 2; c++)
            {
                var s = sheet[r, c].GetStyle();
                s.Bold.Should().Be(true);
                s.Background.Should().Be(Color.Yellow);
            }
    }

    [Fact]
    public void Apply_Pool_Dedupes_The_Style_Across_The_Range()
    {
        // 16 cells, all styled identically — should share ONE NPOI ICellStyle.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.Range("A1:D4").Value(1);
        sheet.Range("A1:D4").Apply(new CellStyle { Bold = true });

        var idx = sheet["A1"].Underlying.CellStyle.Index;
        for (int r = 1; r <= 4; r++)
            for (int c = 1; c <= 4; c++)
                sheet[r, c].Underlying.CellStyle.Index.Should().Be(idx);
    }

    // ---- Merge() ------------------------------------------------------

    [Fact]
    public void Merge_Delegates_To_Sheet_MergeCells()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.Range("A1:C3").Merge();
        sheet.MergedRanges.Should().ContainSingle().Which.Should().Be("A1:C3");
    }

    // ---- Enumeration --------------------------------------------------

    [Fact]
    public void Sparse_Enumeration_Yields_Only_Populated_Cells()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("a");
        sheet["C3"].SetString("c");
        // B2 left empty intentionally.

        var range = sheet.Range("A1:C3");
        var addresses = range.Select(c => c.Address).OrderBy(a => a, StringComparer.Ordinal).ToList();
        addresses.Should().BeEquivalentTo(SparseExpected);
    }

    [Fact]
    public void EnumerateAll_Yields_Every_Coordinate_Materializing_Empties()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("a");

        var addresses = sheet.Range("A1:B2").EnumerateAll().Select(c => c.Address).ToList();
        addresses.Should().BeEquivalentTo(DenseExpected);
    }

    // ---- ClearContents ------------------------------------------------

    [Fact]
    public void ClearContents_Clears_Populated_Cells_But_Preserves_Style_Pool_Index()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet.Range("A1:B2").Value(1).Apply(new CellStyle { NumberFormat = NumberFormats.Currency });
        var styleIdxBefore = sheet["A1"].Underlying.CellStyle.Index;

        sheet.Range("A1:B2").ClearContents();

        for (int r = 1; r <= 2; r++)
            for (int c = 1; c <= 2; c++)
                sheet[r, c].Kind.Should().Be(CellKind.Empty);

        // Style index remains because Clear only clears the cell value.
        sheet["A1"].Underlying.CellStyle.Index.Should().Be(styleIdxBefore);
    }

    // ---- Round-trip ---------------------------------------------------

    [Fact]
    public void Range_Operations_Roundtrip_Through_Save_Open()
    {
        var path = Path.Combine(Path.GetTempPath(), $"range-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("S");
                sheet.Range("A1:C1").Value("header").Apply(new CellStyle { Bold = true });
                sheet.Range("A2:C5").Value(0);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var sheet = wb["S"];
                sheet["B1"].GetString().Should().Be("header");
                sheet["B1"].GetStyle().Bold.Should().Be(true);
                sheet["A5"].GetNumber().Should().Be(0);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
