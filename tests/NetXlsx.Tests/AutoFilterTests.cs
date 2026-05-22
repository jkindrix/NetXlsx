// Coverage for the v1.1 AutoFilter slice: ISheet.SetAutoFilter,
// ClearAutoFilter, HasAutoFilter, AutoFilterRange.

using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class AutoFilterTests
{
    [Fact]
    public void Sheet_Has_No_AutoFilter_By_Default()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.HasAutoFilter.Should().BeFalse();
        sh.AutoFilterRange.Should().BeNull();
    }

    [Fact]
    public void SetAutoFilter_Records_Range()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:D10");
        sh.HasAutoFilter.Should().BeTrue();
        sh.AutoFilterRange.Should().Be("A1:D10");
    }

    [Fact]
    public void SetAutoFilter_Replaces_Existing()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B5");
        sh.SetAutoFilter("C1:E20");
        sh.AutoFilterRange.Should().Be("C1:E20");
    }

    [Fact]
    public void ClearAutoFilter_Removes_The_Filter()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:D10");
        sh.ClearAutoFilter();
        sh.HasAutoFilter.Should().BeFalse();
        sh.AutoFilterRange.Should().BeNull();
    }

    [Fact]
    public void ClearAutoFilter_On_Empty_Sheet_Is_Safe()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.ClearAutoFilter();
        act.Should().NotThrow();
    }

    [Fact]
    public void SetAutoFilter_Rejects_Null_Range()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.SetAutoFilter(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void SetAutoFilter_Rejects_Invalid_Range()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.SetAutoFilter("notarange");
        act.Should().Throw<InvalidCellAddressException>();
    }

    [Fact]
    public void SetAutoFilter_Accepts_Single_Cell_Range()
    {
        // Single-cell autofilter is permitted (Excel treats it as a
        // 1×1 filterable area). Verify we don't reject it.
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1");
        sh.HasAutoFilter.Should().BeTrue();
        // NPOI's CellRangeAddress.FormatAsString collapses 1×1 to "A1".
        sh.AutoFilterRange.Should().Be("A1");
    }

    // ---- Per-column criteria (v1.2 / I-66) ----------------------------

    [Fact]
    public void SetAutoFilterColumn_Requires_AutoFilter_To_Be_Set_First()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.SetAutoFilterColumn(0, FilterCriteria.EqualTo("x"));
        act.Should().Throw<InvalidOperationException>().WithMessage("*SetAutoFilter*");
    }

    [Fact]
    public void SetAutoFilterColumn_Rejects_Negative_Offset()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:C5");
        Action act = () => sh.SetAutoFilterColumn(-1, FilterCriteria.EqualTo("x"));
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void SetAutoFilterColumn_Rejects_Offset_Past_Range_Width()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:C5");   // width 3
        Action act = () => sh.SetAutoFilterColumn(3, FilterCriteria.EqualTo("x"));
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void SetAutoFilterColumn_Rejects_Null_Criteria()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A2");
        Action act = () => sh.SetAutoFilterColumn(0, null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void SetAutoFilterColumn_EqualTo_Persists_Single_CustomFilter()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B5");
        sh.SetAutoFilterColumn(1, FilterCriteria.EqualTo("EU"));

        var af = sh.Underlying.GetCTWorksheet().autoFilter;
        af.filterColumn.Should().HaveCount(1);
        af.filterColumn[0].colId.Should().Be(1u);
        var cf = af.filterColumn[0].customFilters;
        cf.customFilter.Should().HaveCount(1);
        cf.customFilter[0].@operator.Should().Be(NPOI.OpenXmlFormats.Spreadsheet.ST_FilterOperator.equal);
        cf.customFilter[0].val.Should().Be("EU");
        cf.and.Should().BeFalse("single condition — combinator is irrelevant but OOXML default is false");
    }

    [Fact]
    public void SetAutoFilterColumn_Between_Builds_Two_Anded_Conditions()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        sh.SetAutoFilterColumn(0, FilterCriteria.Between(10, 100));

        var cf = sh.Underlying.GetCTWorksheet().autoFilter.filterColumn[0].customFilters;
        cf.customFilter.Should().HaveCount(2);
        cf.and.Should().BeTrue();
        cf.customFilter[0].@operator.Should().Be(NPOI.OpenXmlFormats.Spreadsheet.ST_FilterOperator.greaterThanOrEqual);
        cf.customFilter[0].val.Should().Be("10");
        cf.customFilter[1].@operator.Should().Be(NPOI.OpenXmlFormats.Spreadsheet.ST_FilterOperator.lessThanOrEqual);
        cf.customFilter[1].val.Should().Be("100");
    }

    [Fact]
    public void SetAutoFilterColumn_Or_Combinator_Sets_And_False()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        sh.SetAutoFilterColumn(0,
            FilterCriteria.EqualTo("EU").Or(FilterCriteria.EqualTo("US")));

        var cf = sh.Underlying.GetCTWorksheet().autoFilter.filterColumn[0].customFilters;
        cf.and.Should().BeFalse();
        cf.customFilter.Should().HaveCount(2);
    }

    [Theory]
    [InlineData("Contains")]
    [InlineData("StartsWith")]
    [InlineData("EndsWith")]
    public void SetAutoFilterColumn_String_Pattern_Encodes_With_Wildcards(string kind)
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");

        var criteria = kind switch
        {
            "Contains" => FilterCriteria.Contains("foo"),
            "StartsWith" => FilterCriteria.StartsWith("foo"),
            "EndsWith" => FilterCriteria.EndsWith("foo"),
            _ => throw new ArgumentException("bad kind"),
        };
        sh.SetAutoFilterColumn(0, criteria);

        var val = sh.Underlying.GetCTWorksheet().autoFilter.filterColumn[0].customFilters.customFilter[0].val;
        // Wildcards encode at the appropriate edges.
        switch (kind)
        {
            case "Contains":  val.Should().Be("*foo*"); break;
            case "StartsWith": val.Should().Be("foo*"); break;
            case "EndsWith":   val.Should().Be("*foo"); break;
        }
    }

    [Fact]
    public void SetAutoFilterColumn_Replaces_Existing_Criteria_On_Same_Column()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        sh.SetAutoFilterColumn(0, FilterCriteria.EqualTo("first"));
        sh.SetAutoFilterColumn(0, FilterCriteria.EqualTo("second"));

        var cols = sh.Underlying.GetCTWorksheet().autoFilter.filterColumn;
        cols.Should().HaveCount(1, "replace, not append, when same column re-targets");
        cols[0].customFilters.customFilter[0].val.Should().Be("second");
    }

    [Fact]
    public void ClearAutoFilterColumn_Removes_Just_That_Column()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:C5");
        sh.SetAutoFilterColumn(0, FilterCriteria.EqualTo("a"));
        sh.SetAutoFilterColumn(2, FilterCriteria.EqualTo("c"));

        sh.ClearAutoFilterColumn(0);

        var cols = sh.Underlying.GetCTWorksheet().autoFilter.filterColumn;
        cols.Should().HaveCount(1);
        cols[0].colId.Should().Be(2u);
    }

    [Fact]
    public void ClearAutoFilterColumn_On_Empty_Set_Is_Safe()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        Action act = () => sh.ClearAutoFilterColumn(0);
        act.Should().NotThrow();
    }

    [Fact]
    public void FilterCriteria_And_Then_And_Throws()
    {
        // Excel allows at most two conditions per column.
        Action act = () => FilterCriteria.EqualTo("a")
            .And(FilterCriteria.EqualTo("b"))
            .And(FilterCriteria.EqualTo("c"));
        act.Should().Throw<InvalidOperationException>().WithMessage("*at most two*");
    }

    [Fact]
    public void FilterCriteria_Contains_With_Wildcards_Escapes_Them()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        // Literal "*" and "?" in the search term must be escaped with "~".
        sh.SetAutoFilterColumn(0, FilterCriteria.Contains("a*b?c"));
        var val = sh.Underlying.GetCTWorksheet().autoFilter.filterColumn[0].customFilters.customFilter[0].val;
        val.Should().Be("*a~*b~?c*");
    }

    // ---- FilterCriteria.In(...) (v1.3 / I-68) -------------------------

    [Fact]
    public void FilterCriteria_In_Single_Value_Reduces_To_EqualTo()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        sh.SetAutoFilterColumn(0, FilterCriteria.In("EU"));

        var cf = sh.Underlying.GetCTWorksheet().autoFilter.filterColumn[0].customFilters;
        cf.customFilter.Should().HaveCount(1);
        cf.customFilter[0].@operator.Should().Be(NPOI.OpenXmlFormats.Spreadsheet.ST_FilterOperator.equal);
        cf.customFilter[0].val.Should().Be("EU");
    }

    [Fact]
    public void FilterCriteria_In_Two_Values_Builds_OR_Joined_Equality()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        sh.SetAutoFilterColumn(0, FilterCriteria.In("EU", "US"));

        var cf = sh.Underlying.GetCTWorksheet().autoFilter.filterColumn[0].customFilters;
        cf.customFilter.Should().HaveCount(2);
        cf.and.Should().BeFalse("In() is an OR-of-equality");
        cf.customFilter[0].val.Should().Be("EU");
        cf.customFilter[1].val.Should().Be("US");
    }

    [Fact]
    public void FilterCriteria_In_Three_Or_More_Values_Throws_NotSupported()
    {
        Action act = () => FilterCriteria.In("a", "b", "c");
        act.Should().Throw<NotSupportedException>()
            .WithMessage("*NPOI 2.7.3*")
            .WithMessage("*<filters>*");
    }

    [Fact]
    public void FilterCriteria_In_Rejects_Empty_List()
    {
        Action a = () => FilterCriteria.In();
        a.Should().Throw<ArgumentException>().WithMessage("*at least one value*");
        Action b = () => FilterCriteria.In(System.Linq.Enumerable.Empty<string>());
        b.Should().Throw<ArgumentException>().WithMessage("*at least one value*");
    }

    [Fact]
    public void FilterCriteria_In_Rejects_Null_Entries()
    {
        Action act = () => FilterCriteria.In("EU", null!);
        act.Should().Throw<ArgumentException>().WithMessage("*null entries*");
    }

    [Fact]
    public void FilterCriteria_In_Rejects_Null_Argument()
    {
        Action a = () => FilterCriteria.In((string[])null!);
        a.Should().Throw<ArgumentNullException>();
        Action b = () => FilterCriteria.In((System.Collections.Generic.IEnumerable<string>)null!);
        b.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void FilterCriteria_In_Accepts_IEnumerable_Overload()
    {
        // Verifies the IEnumerable overload — common call shape is
        // `In(myList)` where myList is a List<string> or similar.
        var values = new System.Collections.Generic.List<string> { "A", "B" };
        var criteria = FilterCriteria.In(values);

        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        sh.SetAutoFilterColumn(0, criteria);

        var cf = sh.Underlying.GetCTWorksheet().autoFilter.filterColumn[0].customFilters;
        cf.customFilter.Should().HaveCount(2);
    }

    [Fact]
    public void AutoFilterColumn_Criteria_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"afcol-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh.AppendRow().Set(1, "Region").Set(2, "Revenue");
                sh.AppendRow().Set(1, "EU").Set(2, 100);
                sh.AppendRow().Set(1, "US").Set(2, 200);
                sh.SetAutoFilter("A1:B3");
                sh.SetAutoFilterColumn(0, FilterCriteria.EqualTo("EU").Or(FilterCriteria.EqualTo("US")));
                sh.SetAutoFilterColumn(1, FilterCriteria.GreaterThanOrEqual(50));
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var af = wb["S"].Underlying.GetCTWorksheet().autoFilter;
                af.filterColumn.Should().HaveCount(2);
                af.filterColumn[0].customFilters.customFilter.Should().HaveCount(2);
                af.filterColumn[0].customFilters.and.Should().BeFalse();
                af.filterColumn[1].customFilters.customFilter[0].val.Should().Be("50");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Original test continues below --------------------------------

    [Fact]
    public void AutoFilter_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"af-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh["A1"].SetString("Name");
                sh["B1"].SetString("Score");
                sh["A2"].SetString("Alice"); sh["B2"].SetNumber(90);
                sh["A3"].SetString("Bob");   sh["B3"].SetNumber(85);
                sh.SetAutoFilter("A1:B3");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var sh = wb["S"];
                sh.HasAutoFilter.Should().BeTrue();
                sh.AutoFilterRange.Should().Be("A1:B3");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
