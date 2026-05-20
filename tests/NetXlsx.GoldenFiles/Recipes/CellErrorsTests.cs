// Golden-file test for cookbook recipe 13 (CellErrors).

using System;
using System.IO;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class CellErrorsTests
{
    [Fact]
    public async Task Recipe_Round_Trips_All_Eight_Excel_Error_Codes()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-errors-{Guid.NewGuid():N}.xlsx");
        try
        {
            await CellErrors.Run(path);

            using var wb = Workbook.Open(path);
            var sheet = wb[CellErrors.SheetName];

            // Header row.
            sheet["A1"].GetString().Should().Be("Error code");

            // Each subsequent row's column A is an error cell.
            for (int i = 0; i < CellErrors.Errors.Length; i++)
            {
                var (_, expectedError, _) = CellErrors.Errors[i];
                var cell = sheet[i + 2, 1];   // rows start at 2 (after header)
                cell.Kind.Should().Be(CellKind.Error);
                cell.GetError().Should().Be(expectedError,
                    $"row {i + 2} should classify as {expectedError}");
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
