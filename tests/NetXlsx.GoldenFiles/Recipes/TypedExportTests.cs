// Golden-file test for cookbook recipe 3 (TypedExport).
// Exercises the source-generated extension methods end-to-end.

using System;
using System.IO;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class TypedExportTests
{
    [Fact]
    public async Task Recipe_Writes_Header_Plus_Typed_Rows()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-typed-{Guid.NewGuid():N}.xlsx");
        try
        {
            await TypedExport.Run(path);

            using var wb = Workbook.Open(path);
            var sheet = wb[TypedExport.SheetName];

            // Header (hand-written by the recipe).
            sheet["A1"].GetString().Should().Be("Region");
            sheet["B1"].GetString().Should().Be("Revenue");
            sheet["C1"].GetString().Should().Be("Margin");
            sheet["D1"].GetString().Should().Be("Strategic");

            // Data rows produced by SalesRecord_SheetExtensions.AddRows().
            // Source-declaration order in v0.3.x: Region, Revenue, Margin, Strategic.
            sheet["A2"].GetString().Should().Be("North");
            sheet["B2"].GetNumber().Should().Be(1000.50);
            sheet["C2"].GetNumber().Should().BeApproximately(0.12, 1e-12);
            sheet["D2"].GetBool().Should().Be(true);

            sheet["A3"].GetString().Should().Be("South");
            sheet["B3"].GetNumber().Should().Be(2500.00);
            sheet["D3"].GetBool().Should().Be(false);

            sheet["A4"].GetString().Should().Be("East");
            sheet["B4"].GetNumber().Should().Be(3700.75);
            sheet["D4"].GetBool().Should().Be(true);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task AddRow_Single_Record_Appends_One_Row()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-typed-single-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var writeSheet = wb.AddSheet("Solo");
                writeSheet.AddRow(new SalesRecord
                {
                    Region = "Lone",
                    Revenue = 100m,
                    Margin = 0.5,
                    Strategic = false
                });
                await wb.SaveAsync(path);
            }

            using var read = Workbook.Open(path);
            var sheet = read["Solo"];
            sheet["A1"].GetString().Should().Be("Lone");
            sheet["B1"].GetNumber().Should().Be(100.0);
            sheet["C1"].GetNumber().Should().BeApproximately(0.5, 1e-12);
            sheet["D1"].GetBool().Should().Be(false);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task AddRow_Throws_On_Null_Record()
    {
        // The generator emits ArgumentNullException guards on the
        // record parameter — pinning the contract.
        var path = Path.Combine(Path.GetTempPath(), $"golden-typed-null-{Guid.NewGuid():N}.xlsx");
        try
        {
            using var wb = Workbook.Create();
            var sheet = wb.AddSheet("S");
            Action call = () => sheet.AddRow((SalesRecord)null!);
            call.Should().Throw<ArgumentNullException>();
            await Task.CompletedTask;
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
