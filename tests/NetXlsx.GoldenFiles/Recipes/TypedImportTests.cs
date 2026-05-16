// Golden-file test for cookbook recipe 4 (TypedImport).
// Exercises the source-generator's ReadRows path end-to-end through
// the cookbook recipe.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using FluentAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class TypedImportTests
{
    [Fact]
    public async Task Round_Trips_Records_Verbatim_Via_Source_Generated_ReadRows()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-typed-import-{Guid.NewGuid():N}.xlsx");
        var input = new[]
        {
            new SalesRecord { Region = "Alpha", Revenue = 100m,   Margin = 0.10, Strategic = true  },
            new SalesRecord { Region = "Beta",  Revenue = 250.75m, Margin = 0.20, Strategic = false },
        };

        try
        {
            var parsed = await TypedImport.RunAndReturn(path, input);
            parsed.Should().BeEquivalentTo(input);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task Reads_Empty_Dataset_Cleanly()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-typed-import-empty-{Guid.NewGuid():N}.xlsx");
        try
        {
            var parsed = await TypedImport.RunAndReturn(path, new List<SalesRecord>());
            parsed.Should().BeEmpty();
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task ReadRows_Throws_When_Header_Missing()
    {
        // Write a workbook whose header row is empty, then attempt to
        // ReadRows. The generated body should throw WorkbookException
        // citing the missing header.
        var path = Path.Combine(Path.GetTempPath(), $"golden-typed-import-noheader-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var writeSheet = wb.AddSheet(TypedImport.SheetName);
                // Header row is empty; jump straight to data — but with
                // nothing for the generator to map to, ReadRows should
                // fail loud rather than yield empty.
                writeSheet.AppendRow();    // empty header
                writeSheet.AppendRow().Set(1, "data");
                await wb.SaveAsync(path);
            }

            using var read = await Workbook.OpenAsync(path);
            var sheet = read[TypedImport.SheetName];
            Action act = () => { _ = sheet.ReadRows().ToList(); };
            act.Should().Throw<WorkbookException>()
                .WithMessage("*Header*not found*");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task ReadRows_Throws_NotSupported_For_Headerless_Mode()
    {
        // Decision I-46: header-less ReadRows is deferred to v2; the
        // generated body throws NotSupportedException.
        var path = Path.Combine(Path.GetTempPath(), $"golden-typed-import-headerless-{Guid.NewGuid():N}.xlsx");
        try
        {
            await TypedExport.Run(path);
            using var read = await Workbook.OpenAsync(path);
            var sheet = read[TypedExport.SheetName];

            Action act = () => { _ = sheet.ReadRows(headerRow: null).ToList(); };
            act.Should().Throw<NotSupportedException>()
                .WithMessage("*Header-less*deferred*");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
