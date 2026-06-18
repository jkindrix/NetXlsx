// Golden-file smoke tests for the eleven post-v1.1 cookbook recipes
// (v1.2–v2.0 surface, R-23). Each test runs the recipe end-to-end and
// asserts the salient feature output — either via the public read-back
// API where one exists (conditional-format count, sorted cell values,
// picture border, IsMacroEnabled) or, where the surface has no live
// read-back, via the persisted OOXML part (charts, panes, outline
// levels, connectors, autofilter columns, totals row) using the shared
// SavedOoxml inspector. The OOXML bytes are the real contract.

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using AwesomeAssertions;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using NetXlsx.Tests;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class PostV11RecipeSmoke
{
    private static async Task<string> RunRecipe(Func<string, Task> recipe, string tag, string ext = "xlsx")
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-v12-{tag}-{Guid.NewGuid():N}.{ext}");
        await recipe(path);
        return path;
    }

    [Fact]
    public async Task ConditionalFormatting_Roundtrips_With_Three_Rules()
    {
        var path = await RunRecipe(ConditionalFormatting.Run, "cf");
        try
        {
            using var wb = Workbook.Open(path);
            // All three rules (cell-value, formula, color-scale) target the
            // same range and rehydrate on open.
            wb[ConditionalFormatting.SheetName].ConditionalFormattingCount.Should().Be(3);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task Charts_Emit_Two_Graphic_Frames_In_The_Drawing()
    {
        var path = await RunRecipe(Charts.Run, "chart");
        try
        {
            // Each chart survives as a graphic frame in the drawing part.
            SavedOoxml.PartFromFile(path, "xl/drawings/drawing1.xml")
                .Descendants(SavedOoxml.Xdr + "graphicFrame")
                .Should().HaveCount(2);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task SortingRanges_Sorts_By_Dept_Then_Salary_Descending()
    {
        var path = await RunRecipe(SortingRanges.Run, "sort");
        try
        {
            using var wb = Workbook.Open(path);
            var sh = wb[SortingRanges.SheetName];
            // Header untouched; Eng before Sales, salary descending within.
            sh["B1"].GetString().Should().Be("Dept");
            sh["A2"].GetString().Should().Be("Cleo");   // Eng, 165k
            sh["A3"].GetString().Should().Be("Ada");    // Eng, 140k
            sh["A4"].GetString().Should().Be("Dane");   // Sales, 120k
            sh["A5"].GetString().Should().Be("Bert");   // Sales, 90k
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task SplitAndFreezePanes_Emit_Frozen_And_Split_Panes()
    {
        var path = await RunRecipe(SplitAndFreezePanes.Run, "pane");
        try
        {
            var frozen = Pane(SavedOoxml.PartFromFile(path, "xl/worksheets/sheet1.xml"));
            frozen.Should().NotBeNull();
            ((string?)frozen!.Attribute("state")).Should().Be("frozen");

            var split = Pane(SavedOoxml.PartFromFile(path, "xl/worksheets/sheet2.xml"));
            split.Should().NotBeNull();
            ((string?)split!.Attribute("state")).Should().NotBe("frozen");
            ((double?)split.Attribute("xSplit")).Should().Be(2000);
            ((double?)split.Attribute("ySplit")).Should().Be(3000);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task RowAndColumnGrouping_Emits_Nested_Outline_Levels()
    {
        var path = await RunRecipe(RowAndColumnGrouping.Run, "group");
        try
        {
            var sheet = SavedOoxml.PartFromFile(path, "xl/worksheets/sheet1.xml");
            // Rows 2–3 are nested one level deeper than row 4.
            OutlineLevel(sheet, 2).Should().Be(2);
            OutlineLevel(sheet, 3).Should().Be(2);
            OutlineLevel(sheet, 4).Should().Be(1);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task ShapesAndConnectors_Emit_Two_Shapes_And_One_Connector()
    {
        var path = await RunRecipe(ShapesAndConnectors.Run, "shape");
        try
        {
            var drawing = SavedOoxml.PartFromFile(path, "xl/drawings/drawing1.xml");
            drawing.Descendants(SavedOoxml.Xdr + "sp").Should().HaveCount(2);
            drawing.Descendants(SavedOoxml.Xdr + "cxnSp").Should().HaveCount(1);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task PictureBorders_Roundtrips_With_Border_On_Reopen()
    {
        var path = await RunRecipe(PictureBorders.Run, "picborder");
        try
        {
            using var wb = Workbook.Open(path);
            var pic = wb[PictureBorders.SheetName].Pictures.Single();
            pic.Border.Should().NotBeNull();
            pic.Border!.WidthPoints.Should().BeApproximately(PictureBorders.BorderWidthPoints, 0.01);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task TotalsRows_Emit_Sum_And_Average_Functions()
    {
        var path = await RunRecipe(TotalsRows.Run, "totals");
        try
        {
            var table = SavedOoxml.PartFromFile(path, "xl/tables/table1.xml").Root!;
            ((string?)table.Attribute("totalsRowCount")).Should().Be("1");
            var cols = table.Element(SavedOoxml.Main + "tableColumns")!
                .Elements(SavedOoxml.Main + "tableColumn").ToList();
            ((string?)cols[1].Attribute("totalsRowFunction")).Should().Be("sum");
            ((string?)cols[2].Attribute("totalsRowFunction")).Should().Be("average");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task AutoFilterCriteria_Emit_Two_Filtered_Columns()
    {
        var path = await RunRecipe(AutoFilterCriteria.Run, "af");
        try
        {
            using var wb = Workbook.Open(path);
            wb[AutoFilterCriteria.SheetName].HasAutoFilter.Should().BeTrue();

            var filterCols = SavedOoxml.PartFromFile(path, "xl/worksheets/sheet1.xml")
                .Descendants(SavedOoxml.Main + "filterColumn").ToList();
            // One filterColumn for Region (Or pair) and one for Amount (>=).
            filterCols.Select(c => (uint?)c.Attribute("colId"))
                .Should().BeEquivalentTo(new uint?[] { 0u, 2u });
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task NamedStyleIntegration_Persists_Name_To_Ooxml_And_Shares_StylePool_Entry()
    {
        var path = await RunRecipe(NamedStyleIntegration.Run, "named");
        try
        {
            using var wb = Workbook.Open(path);
            var sh = wb[NamedStyleIntegration.SheetName];
            // Both headings resolve to the same cellXfs index and read back bold.
            sh["A1"].GetStyle().Bold.Should().Be(true);
            sh["A3"].GetStyle().Bold.Should().Be(true);
            var sheetXml = SavedOoxml.PartFromFile(path, "xl/worksheets/sheet1.xml");
            SavedOoxml.CellStyleIndex(sheetXml, "A1")
                .Should().Be(SavedOoxml.CellStyleIndex(sheetXml, "A3"),
                    "the named style dedups to a single shared style-pool entry");
            // I-67: the name round-trips through the OOXML cellStyles table,
            // so Workbook.Open rehydrates the name -> style map.
            wb.RegisteredStyleNames.Should().Contain(NamedStyleIntegration.HeadingStyle);
            wb.GetRegisteredStyle(NamedStyleIntegration.HeadingStyle)!.Bold.Should().Be(true);
            // The <cellStyle name="Heading"> entry is present in styles.xml.
            SavedOoxml.PartFromFile(path, "xl/styles.xml")
                .Descendants(SavedOoxml.Main + "cellStyle")
                .Select(e => (string?)e.Attribute("name"))
                .Should().Contain(NamedStyleIntegration.HeadingStyle);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task MacroEnabledPassthrough_Stays_MacroEnabled_Across_RoundTrip()
    {
        var path = await RunRecipe(MacroEnabledPassthrough.Run, "xlsm", ext: "xlsm");
        try
        {
            using var wb = Workbook.Open(path);
            wb.IsMacroEnabled.Should().BeTrue();
            wb[MacroEnabledPassthrough.SheetName]["A1"].GetString().Should().Be("Metric");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- helpers (mirror SplitPaneTests / GroupingTests) ----

    private static XElement? Pane(XDocument sheetXml)
        => sheetXml.Root!
            .Element(SavedOoxml.Main + "sheetViews")?
            .Element(SavedOoxml.Main + "sheetView")?
            .Element(SavedOoxml.Main + "pane");

    private static int OutlineLevel(XDocument sheetXml, int rowNumber)
        => (int?)sheetXml.Root!.Element(SavedOoxml.Main + "sheetData")!
            .Elements(SavedOoxml.Main + "row")
            .FirstOrDefault(r => (int?)r.Attribute("r") == rowNumber)?
            .Attribute("outlineLevel") ?? 0;
}
