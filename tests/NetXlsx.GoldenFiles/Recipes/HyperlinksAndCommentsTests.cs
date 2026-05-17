// Golden-file test for cookbook recipe 8 (HyperlinksAndComments).

using System;
using System.IO;
using System.Threading.Tasks;
using NetXlsx;
using NetXlsx.Cookbook.Recipes;
using FluentAssertions;
using Xunit;

namespace NetXlsx.GoldenFiles.Recipes;

public class HyperlinksAndCommentsTests
{
    [Fact]
    public async Task Recipe_Writes_All_Four_Hyperlink_Schemes_And_Two_Comment_Authors()
    {
        var path = Path.Combine(Path.GetTempPath(), $"golden-hclinks-{Guid.NewGuid():N}.xlsx");
        try
        {
            await HyperlinksAndComments.Run(path);

            using var wb = Workbook.Open(path);
            wb.SheetCount.Should().Be(2);

            var notes = wb[HyperlinksAndComments.SheetName];

            // Header.
            notes["A1"].GetString().Should().Be("Item");
            notes["B1"].GetString().Should().Be("Link");

            // Four hyperlink rows.
            notes["B3"].GetHyperlink().Should().Be("https://example.com/releases");
            notes["B3"].GetString().Should().Be("Release page");

            notes["B4"].GetHyperlink().Should().Be("mailto:maintainer@example.com");
            // No display string passed and the cell was empty → falls
            // back to the raw target.
            notes["B4"].GetString().Should().Be("mailto:maintainer@example.com");

            notes["B5"].GetHyperlink().Should().Be("file:///net/share/releases/1.0.0/notes.pdf");
            notes["B5"].GetString().Should().Be("1.0.0/notes.pdf");

            // Internal #Sheet!Range — NPOI stores Document-type body
            // without the leading '#'.
            notes["B6"].GetHyperlink().Should().Be("Changelog!A2");
            notes["B6"].GetString().Should().Be("Changelog row");

            // Three comments — two default authors, one explicit.
            notes["A3"].GetComment().Should().Be("Always link the canonical release page.");
            notes["A3"].GetCommentAuthor().Should().Be("NetXlsx");   // I11 default

            notes["A4"].GetComment().Should().Be("Distribution list, not a single human.");
            notes["A4"].GetCommentAuthor().Should().Be("release-bot");  // explicit override

            notes["A5"].GetComment().Should().Be("Path is canonical on the build server.");
            notes["A5"].GetCommentAuthor().Should().Be("NetXlsx");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
