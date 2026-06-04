// Coverage for v1.0-A WorkbookOptions wiring on the random-access side:
// entry-point overloads + write-side limit enforcement +
// DefaultFontName/DefaultFontSize.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class WorkbookOptionsWritePathTests
{
    // ---- Entry-point overloads accept WorkbookOptions ----------------

    [Fact]
    public void Create_With_Null_Options_Is_Equivalent_To_Default_Options()
    {
        using var wb = Workbook.Create(null);
        // Default sheet count == 0, no observable misbehavior.
        wb.SheetCount.Should().Be(0);
    }

    [Fact]
    public void Create_Default_Has_Excel_Hard_Cap_Behavior()
    {
        // With no options passed, behavior matches the pre-options
        // contract — MaxRowsPerSheet defaults to Excel's hard cap so
        // nothing rejects past prior bounds.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        // Indexer at row 1, col 1 works as before.
        sheet[1, 1].Address.Should().Be("A1");
    }

    [Fact]
    public void Open_With_Options_Round_Trips_Options_Through_File()
    {
        var path = Path.Combine(Path.GetTempPath(), $"opts-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create(new WorkbookOptions { DefaultFontName = "Arial" }))
            {
                wb.AddSheet("S")[1, 1].SetString("x");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path, new WorkbookOptions { MaxCellTextLength = 100 }))
            {
                // Open didn't reject the file. Now write with the tighter
                // limit on the opened workbook.
                var sheet = wb["S"];
                ((Action)(() => sheet[2, 1].SetString(new string('x', 101))))
                    .Should().Throw<ResourceLimitExceededException>();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- MaxCellTextLength ------------------------------------------

    [Fact]
    public void Default_MaxCellTextLength_Is_Excel_Hard_Cap_32767()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        // 32,767 chars: OK. 32,768: throws.
        sheet[1, 1].SetString(new string('x', 32_767));
        ((Action)(() => sheet[1, 2].SetString(new string('x', 32_768))))
            .Should().Throw<ResourceLimitExceededException>()
            .Where(ex => ex.Limit == 32_767 && ex.Actual == 32_768);
    }

    [Fact]
    public void Configured_MaxCellTextLength_Enforced_Below_Excel_Hard_Cap()
    {
        using var wb = Workbook.Create(new WorkbookOptions { MaxCellTextLength = 50 });
        var sheet = wb.AddSheet("S");
        sheet[1, 1].SetString(new string('a', 50));
        ((Action)(() => sheet[1, 2].SetString(new string('a', 51))))
            .Should().Throw<ResourceLimitExceededException>()
            .Where(ex => ex.LimitName == "cell text length" && ex.Limit == 50 && ex.Actual == 51);
    }

    // ---- MaxRowsPerSheet --------------------------------------------

    [Fact]
    public void Default_MaxRowsPerSheet_Allows_Beyond_Configured_Lower_Cap()
    {
        // Default == Excel hard cap; with no override, row 1,048,576 is
        // valid. (We only test the index-validation gate, not actually
        // materializing 1M rows.)
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        ((Action)(() => { _ = sheet[CellAddress.MaxRow, 1].Address; }))
            .Should().NotThrow();
        ((Action)(() => { _ = sheet[CellAddress.MaxRow + 1, 1]; })).Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Configured_MaxRowsPerSheet_Caps_AppendRow_And_Indexer()
    {
        using var wb = Workbook.Create(new WorkbookOptions { MaxRowsPerSheet = 5 });
        var sheet = wb.AddSheet("S");
        for (int i = 0; i < 5; i++) sheet.AppendRow();
        ((Action)(() => sheet.AppendRow())).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*exceed the configured row limit of 5*");

        ((Action)(() => { _ = sheet[6, 1]; })).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*row must be in [1, 5]*");

        ((Action)(() => sheet.Row(6))).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*row index must be in [1, 5]*");
    }

    // ---- MaxColsPerSheet --------------------------------------------

    [Fact]
    public void Configured_MaxColsPerSheet_Caps_Indexer_And_Row_Cell_And_Column_Factory()
    {
        using var wb = Workbook.Create(new WorkbookOptions { MaxColsPerSheet = 3 });
        var sheet = wb.AddSheet("S");
        sheet[1, 3].SetString("ok-at-cap");

        ((Action)(() => { _ = sheet[1, 4]; })).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*column must be in [1, 3]*");

        var row = sheet.AppendRow();
        ((Action)(() => row.Cell(4))).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*column must be in [1, 3]*");

        ((Action)(() => sheet.Column(4))).Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*column index must be in [1, 3]*");
    }

    // ---- DefaultFontName / DefaultFontSize --------------------------

    [Fact]
    public void DefaultFontName_Is_Applied_To_Workbook_Default_Font()
    {
        using var wb = Workbook.Create(new WorkbookOptions
        {
            DefaultFontName = "Arial",
            DefaultFontSize = 14,
        });
        var (name, size) = Font0(wb);
        name.Should().Be("Arial");
        size.Should().Be(14);
    }

    [Fact]
    public void DefaultFontName_Default_Is_Calibri_11()
    {
        using var wb = Workbook.Create();
        var (name, size) = Font0(wb);
        name.Should().Be("Calibri");
        size.Should().Be(11);
    }

    // ---- Round-trip: defaults are still applied after open ----------

    [Fact]
    public void Open_Preserves_File_Default_Font_Regardless_Of_Options()
    {
        // Opening a file never overwrites its default font (font 0). The
        // options' DefaultFontName/Size are write-side defaults applied
        // only when creating a new workbook — silently swapping a file's
        // default font on Open would mis-render every cell that inherits
        // it (e.g. a modern Excel file's "Aptos Narrow" becoming Calibri).
        var path = Path.Combine(Path.GetTempPath(), $"opts-font-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create(new WorkbookOptions { DefaultFontName = "Verdana", DefaultFontSize = 12 }))
            {
                wb.AddSheet("S")[1, 1].SetString("authored-with-verdana");
                wb.Save(path);
            }
            // Reopen with NO options — the file's "Verdana" must survive,
            // not get clobbered with NetXlsx's Calibri/11 default.
            using (var wb = Workbook.Open(path))
            {
                var (name, size) = Font0(wb);
                name.Should().Be("Verdana");
                size.Should().Be(12);
            }
            // Reopen with DIFFERENT options — the file still wins.
            using (var wb = Workbook.Open(path, new WorkbookOptions { DefaultFontName = "Arial", DefaultFontSize = 14 }))
            {
                var (name, size) = Font0(wb);
                name.Should().Be("Verdana");
                size.Should().Be(12);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- helpers ------------------------------------------------------

    /// <summary>Font 0 (the workbook default) from the persisted stylesheet.</summary>
    private static (string Name, double Size) Font0(IWorkbook wb)
    {
        var font = SavedOoxml.StylesXml(wb).Root!
            .Element(SavedOoxml.Main + "fonts")!
            .Elements(SavedOoxml.Main + "font").First();
        return (
            (string)font.Element(SavedOoxml.Main + "name")!.Attribute("val")!,
            (double)font.Element(SavedOoxml.Main + "sz")!.Attribute("val")!);
    }
}
