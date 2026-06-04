// Golden-file smoke tests for the seven v1.1 cookbook recipes.
// Each test runs the recipe end-to-end, re-opens the output via
// Workbook.Open, and asserts the salient v1.1-feature output.

using System;
using System.IO;
using System.Threading.Tasks;
using AwesomeAssertions;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class V11RecipeSmoke
{
    private static async Task<string> RunRecipe(Func<string, Task> recipe, string tag)
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-v11-{tag}-{Guid.NewGuid():N}.xlsx");
        await recipe(path);
        return path;
    }

    [Fact]
    public async Task RichTextCells_Roundtrips_With_Formatting_Runs()
    {
        var path = await RunRecipe(RichTextCells.Run, "rt");
        try
        {
            using var wb = Workbook.Open(path);
            var sh = wb[RichTextCells.SheetName];
            sh["A1"].GetRichText().Should().NotBeNull("A1 was set via SetRichText");
            sh["A1"].GetString().Should().StartWith("Release: 1.1.0");
            sh["A2"].GetRichText().Should().NotBeNull();
            sh["A3"].GetRichText().Should().NotBeNull();
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task ExcelTables_Roundtrips_With_Table_And_AutoFilter()
    {
        var path = await RunRecipe(ExcelTables.Run, "tbl");
        try
        {
            using var wb = Workbook.Open(path);
            var sales = wb[ExcelTables.SheetName];
            sales.Tables.Should().HaveCount(1);
            sales.Tables[0].Name.Should().Be("QuarterlySales");
            sales.Tables[0].StyleName.Should().Be("TableStyleMedium2");

            var inv = wb["Inventory"];
            inv.HasAutoFilter.Should().BeTrue();
            inv.AutoFilterRange.Should().Be("A1:B4");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task EmbeddedImages_Roundtrips_With_Two_Pictures()
    {
        var path = await RunRecipe(EmbeddedImages.Run, "img");
        try
        {
            using var wb = Workbook.Open(path);
            wb[EmbeddedImages.SheetName].Pictures.Should().HaveCount(2);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task ProtectedTemplate_Roundtrips_With_Sheet_And_Workbook_Locks()
    {
        var path = await RunRecipe(ProtectedTemplate.Run, "prot");
        try
        {
            using var wb = Workbook.Open(path);
            wb.IsProtected.Should().BeTrue("workbook structure locked");
            wb["Reference"].IsProtected.Should().BeTrue();
            wb[ProtectedTemplate.SheetName].IsProtected.Should().BeTrue();
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task ValidatedInputForm_Roundtrips_With_All_Five_Validations()
    {
        var path = await RunRecipe(ValidatedInputForm.Run, "val");
        try
        {
            using var wb = Workbook.Open(path);
            NetXlsx.Tests.SavedOoxml.PartFromFile(path, "xl/worksheets/sheet1.xml").Root!
                .Element(NetXlsx.Tests.SavedOoxml.Main + "dataValidations")!
                .Elements(NetXlsx.Tests.SavedOoxml.Main + "dataValidation")
                .Should().HaveCount(5);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task BrandedStyles_Roundtrips_With_Applied_Styles()
    {
        var path = await RunRecipe(BrandedStyles.Run, "brand");
        try
        {
            using var wb = Workbook.Open(path);
            var sh = wb[BrandedStyles.SheetName];
            // Named-style map is not rehydrated (decision I-57), but the
            // per-cell visual style is preserved through the style-pool
            // dedup. Spot-check the header row.
            var h = sh["A1"].GetStyle();
            h.Bold.Should().Be(true);
            h.FontSize.Should().Be(14);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public async Task CustomListConverter_Roundtrips_With_Encoded_Tags()
    {
        var path = await RunRecipe(CustomListConverter.Run, "conv");
        try
        {
            using var wb = Workbook.Open(path);
            var sh = wb[CustomListConverter.SheetName];
            sh["C2"].GetString().Should().Be("security;open-path;v1.1");
            sh["C3"].GetString().Should().Be("perf;bench");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
