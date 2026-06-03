// I-82 engine swap — CF/validation/tables/autofilter/sort slice: AutoFilter
// conformance.
//
// Mirrors the NPOI engine's AutoFilterTests contract on the Open XML SDK
// engine (decisions I-56 + I-66): set/clear/replace semantics, per-column
// custom-filter criteria (operators, AND/OR combinators, wildcard string
// patterns), offset validation, and file round-trips. Where the NPOI tests
// reach through sh.Underlying.GetCTWorksheet(), these assert against the
// live SDK DOM via IWorkbook.OpenXmlDocument.
//
// Oracle-pinned extras (SDK-quirk #11 habit — facts confirmed by dumping the
// NPOI engine's XML): SetAutoFilter creates/updates the hidden built-in
// _xlnm._FilterDatabase defined name (no 1x1 collapse, quoted-when-needed
// sheet name); re-setting the range keeps existing filterColumn entries;
// ClearAutoFilter leaves the stale name in place.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.OoxmlEngine.Tests;

public class AutoFilterTests
{
    private static S.AutoFilter? AutoFilterOf(IWorkbook wb)
        => wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single()
            .Worksheet!.GetFirstChild<S.AutoFilter>();

    [Fact]
    public void Sheet_Has_No_AutoFilter_By_Default()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.HasAutoFilter.Should().BeFalse();
        sh.AutoFilterRange.Should().BeNull();
    }

    [Fact]
    public void SetAutoFilter_Records_Range()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:D10");
        sh.HasAutoFilter.Should().BeTrue();
        sh.AutoFilterRange.Should().Be("A1:D10");
    }

    [Fact]
    public void SetAutoFilter_Replaces_Existing()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B2");
        sh.SetAutoFilter("A1:C3");
        sh.AutoFilterRange.Should().Be("A1:C3");
        // Only one <autoFilter> element exists.
        wb.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single()
            .Worksheet!.Elements<S.AutoFilter>().Should().HaveCount(1);
    }

    [Fact]
    public void SetAutoFilter_Replace_Keeps_Existing_Column_Criteria()
    {
        // Oracle-pinned: NPOI's SetAutoFilter replaces only @ref — existing
        // <filterColumn> entries survive a range re-set.
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B2");
        sh.SetAutoFilterColumn(0, FilterCriteria.EqualTo("EU"));
        sh.SetAutoFilter("A1:C3");

        var af = AutoFilterOf(wb)!;
        af.Reference!.Value.Should().Be("A1:C3");
        af.Elements<S.FilterColumn>().Should().HaveCount(1);
    }

    [Fact]
    public void ClearAutoFilter_Removes_The_Filter()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B2");
        sh.ClearAutoFilter();
        sh.HasAutoFilter.Should().BeFalse();
        sh.AutoFilterRange.Should().BeNull();
    }

    [Fact]
    public void ClearAutoFilter_On_Empty_Sheet_Is_Safe()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        Action act = () => sh.ClearAutoFilter();
        act.Should().NotThrow();
    }

    [Fact]
    public void SetAutoFilter_Rejects_Null_Range()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        Action act = () => sh.SetAutoFilter(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void SetAutoFilter_Rejects_Invalid_Range()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        Action act = () => sh.SetAutoFilter("NOT_A_RANGE");
        act.Should().Throw<InvalidCellAddressException>();
    }

    [Fact]
    public void SetAutoFilter_Accepts_Single_Cell_Range()
    {
        // Single-cell autofilter is permitted (Excel treats it as a
        // 1×1 filterable area); the range collapses to "A1" (NPOI's
        // CellRangeAddress.FormatAsString parity).
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1");
        sh.HasAutoFilter.Should().BeTrue();
        sh.AutoFilterRange.Should().Be("A1");
    }

    // ---- the hidden _xlnm._FilterDatabase built-in name ----------------

    [Fact]
    public void SetAutoFilter_Creates_The_FilterDatabase_Name()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B2");

        var name = wb.NamedRanges.Should().ContainSingle().Subject;
        name.Name.Should().Be("_xlnm._FilterDatabase");
        name.Formula.Should().Be("S!$A$1:$B$2");
        name.SheetScope.Should().Be("S");
        // hidden="1", like Excel and NPOI emit it.
        wb.OpenXmlDocument!.WorkbookPart!.Workbook!
            .GetFirstChild<S.DefinedNames>()!.Elements<S.DefinedName>()
            .Single().Hidden!.Value.Should().BeTrue();
    }

    [Fact]
    public void SetAutoFilter_Updates_The_FilterDatabase_Name_On_Replace()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B2");
        sh.SetAutoFilter("A1:C3");

        var name = wb.NamedRanges.Should().ContainSingle().Subject;
        name.Formula.Should().Be("S!$A$1:$C$3");
    }

    [Fact]
    public void SetAutoFilter_SingleCell_Name_Does_Not_Collapse()
    {
        // Oracle-pinned: the defined name always carries the r1:r2 form even
        // for a 1×1 filter ('S'-quoted when the sheet name needs it).
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("My Sheet");
        sh.SetAutoFilter("A1");

        wb.NamedRanges.Single().Formula.Should().Be("'My Sheet'!$A$1:$A$1");
    }

    [Fact]
    public void ClearAutoFilter_Leaves_The_FilterDatabase_Name()
    {
        // NPOI parity: the stale name stays; Excel tolerates it.
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B2");
        sh.ClearAutoFilter();

        wb.NamedRanges.Should().ContainSingle()
            .Which.Name.Should().Be("_xlnm._FilterDatabase");
    }

    // ---- Per-column criteria (decision I-66) ----------------------------

    [Fact]
    public void SetAutoFilterColumn_Requires_AutoFilter_To_Be_Set_First()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        Action act = () => sh.SetAutoFilterColumn(0, FilterCriteria.EqualTo("x"));
        act.Should().Throw<InvalidOperationException>();
    }

    [Fact]
    public void SetAutoFilterColumn_Rejects_Negative_Offset()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B5");
        Action act = () => sh.SetAutoFilterColumn(-1, FilterCriteria.EqualTo("x"));
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void SetAutoFilterColumn_Rejects_Offset_Past_Range_Width()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B5");
        Action act = () => sh.SetAutoFilterColumn(2, FilterCriteria.EqualTo("x"));
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void SetAutoFilterColumn_Rejects_Null_Criteria()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A2");
        Action act = () => sh.SetAutoFilterColumn(0, null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void SetAutoFilterColumn_EqualTo_Persists_Single_CustomFilter()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:B5");
        sh.SetAutoFilterColumn(1, FilterCriteria.EqualTo("EU"));

        var af = AutoFilterOf(wb)!;
        var cols = af.Elements<S.FilterColumn>().ToList();
        cols.Should().HaveCount(1);
        cols[0].ColumnId!.Value.Should().Be(1u);
        var cf = cols[0].GetFirstChild<S.CustomFilters>()!;
        var conditions = cf.Elements<S.CustomFilter>().ToList();
        conditions.Should().HaveCount(1);
        conditions[0].Operator!.Value.Should().Be(S.FilterOperatorValues.Equal);
        conditions[0].Val!.Value.Should().Be("EU");
        (cf.And?.Value ?? false).Should().BeFalse("single condition — combinator is irrelevant but OOXML default is false");
    }

    [Fact]
    public void SetAutoFilterColumn_Between_Builds_Two_Anded_Conditions()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        sh.SetAutoFilterColumn(0, FilterCriteria.Between(10, 100));

        var cf = AutoFilterOf(wb)!.Elements<S.FilterColumn>().Single().GetFirstChild<S.CustomFilters>()!;
        var conditions = cf.Elements<S.CustomFilter>().ToList();
        conditions.Should().HaveCount(2);
        cf.And!.Value.Should().BeTrue();
        conditions[0].Operator!.Value.Should().Be(S.FilterOperatorValues.GreaterThanOrEqual);
        conditions[0].Val!.Value.Should().Be("10");
        conditions[1].Operator!.Value.Should().Be(S.FilterOperatorValues.LessThanOrEqual);
        conditions[1].Val!.Value.Should().Be("100");
    }

    [Fact]
    public void SetAutoFilterColumn_Or_Combinator_Sets_And_False()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        sh.SetAutoFilterColumn(0,
            FilterCriteria.EqualTo("EU").Or(FilterCriteria.EqualTo("US")));

        var cf = AutoFilterOf(wb)!.Elements<S.FilterColumn>().Single().GetFirstChild<S.CustomFilters>()!;
        (cf.And?.Value ?? false).Should().BeFalse();
        cf.Elements<S.CustomFilter>().Should().HaveCount(2);
    }

    [Theory]
    [InlineData("Contains")]
    [InlineData("StartsWith")]
    [InlineData("EndsWith")]
    public void SetAutoFilterColumn_String_Pattern_Encodes_With_Wildcards(string kind)
    {
        using var wb = Workbook.CreateOoxml();
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

        var val = AutoFilterOf(wb)!.Elements<S.FilterColumn>().Single()
            .GetFirstChild<S.CustomFilters>()!.Elements<S.CustomFilter>().Single().Val!.Value;
        switch (kind)
        {
            case "Contains": val.Should().Be("*foo*"); break;
            case "StartsWith": val.Should().Be("foo*"); break;
            case "EndsWith": val.Should().Be("*foo"); break;
        }
    }

    [Fact]
    public void SetAutoFilterColumn_Replaces_Existing_Criteria_On_Same_Column()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        sh.SetAutoFilterColumn(0, FilterCriteria.EqualTo("first"));
        sh.SetAutoFilterColumn(0, FilterCriteria.EqualTo("second"));

        var cols = AutoFilterOf(wb)!.Elements<S.FilterColumn>().ToList();
        cols.Should().HaveCount(1, "replace, not append, when same column re-targets");
        cols[0].GetFirstChild<S.CustomFilters>()!.Elements<S.CustomFilter>()
            .Single().Val!.Value.Should().Be("second");
    }

    [Fact]
    public void ClearAutoFilterColumn_Removes_Just_That_Column()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:C5");
        sh.SetAutoFilterColumn(0, FilterCriteria.EqualTo("a"));
        sh.SetAutoFilterColumn(2, FilterCriteria.EqualTo("c"));

        sh.ClearAutoFilterColumn(0);

        var cols = AutoFilterOf(wb)!.Elements<S.FilterColumn>().ToList();
        cols.Should().HaveCount(1);
        cols[0].ColumnId!.Value.Should().Be(2u);
    }

    [Fact]
    public void ClearAutoFilterColumn_On_Empty_Set_Is_Safe()
    {
        using var wb = Workbook.CreateOoxml();
        var sh = wb.AddSheet("S");
        sh.SetAutoFilter("A1:A5");
        Action act = () => sh.ClearAutoFilterColumn(0);
        act.Should().NotThrow();
        // And with no AutoFilter at all:
        sh.ClearAutoFilter();
        act.Should().NotThrow();
    }

    // ---- round-trips ----------------------------------------------------

    [Fact]
    public void AutoFilter_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-af-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var sh = wb.AddSheet("S");
                sh["A1"].SetString("Name");
                sh["B1"].SetString("Score");
                sh["A2"].SetString("Alice"); sh["B2"].SetNumber(90);
                sh["A3"].SetString("Bob"); sh["B3"].SetNumber(85);
                sh.SetAutoFilter("A1:B3");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                var sh = wb["S"];
                sh.HasAutoFilter.Should().BeTrue();
                sh.AutoFilterRange.Should().Be("A1:B3");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void AutoFilterColumn_Criteria_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-afcol-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.CreateOoxml())
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
            using (var wb = Workbook.OpenOoxml(path))
            {
                var af = AutoFilterOf(wb)!;
                var cols = af.Elements<S.FilterColumn>().ToList();
                cols.Should().HaveCount(2);
                var first = cols[0].GetFirstChild<S.CustomFilters>()!;
                first.Elements<S.CustomFilter>().Should().HaveCount(2);
                (first.And?.Value ?? false).Should().BeFalse();
                cols[1].GetFirstChild<S.CustomFilters>()!.Elements<S.CustomFilter>()
                    .First().Val!.Value.Should().Be("50");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
