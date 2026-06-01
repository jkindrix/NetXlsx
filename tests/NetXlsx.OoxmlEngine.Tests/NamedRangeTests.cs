// I-82 engine swap — structure slice: named-range conformance.
//
// Mirrors the NPOI engine's IWorkbook.AddNamedRange / NamedRanges / INamedRange
// contract on the Open XML SDK engine: workbook- and sheet-scoped names, the
// leading-'=' strip, case-insensitive workbook-wide uniqueness, Excel name-rule
// validation, sheetScope resolution + SheetNameException, and Save/Open
// round-trip of names + scope through real <definedNames> OOXML.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class NamedRangeTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-names-{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void NamedRanges_Is_Empty_Initially()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S");
        wb.NamedRanges.Should().BeEmpty();
    }

    [Fact]
    public void AddNamedRange_WorkbookScoped()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Sheet1");
        var nr = wb.AddNamedRange("Sales", "Sheet1!$A$1:$B$10");

        nr.Name.Should().Be("Sales");
        nr.Formula.Should().Be("Sheet1!$A$1:$B$10");
        nr.SheetScope.Should().BeNull();

        wb.NamedRanges.Should().ContainSingle().Which.Name.Should().Be("Sales");
    }

    [Fact]
    public void AddNamedRange_SheetScoped_Resolves_The_Scope()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("First");
        wb.AddSheet("Second");
        var nr = wb.AddNamedRange("Local", "Second!$A$1", sheetScope: "Second");

        nr.SheetScope.Should().Be("Second");
        // localSheetId is the 0-based document-order index — "Second" is index 1.
        wb.NamedRanges.Single().SheetScope.Should().Be("Second");
    }

    [Fact]
    public void AddNamedRange_Strips_A_Leading_Equals()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S");
        var nr = wb.AddNamedRange("Body", "=S!$A$1");
        nr.Formula.Should().Be("S!$A$1");
    }

    [Fact]
    public void AddNamedRange_Rejects_A_Duplicate_Name_CaseInsensitively()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S");
        wb.AddNamedRange("Total", "S!$A$1");

        var act = () => wb.AddNamedRange("TOTAL", "S!$B$1");
        act.Should().Throw<ArgumentException>().WithMessage("*already exists*");
        wb.NamedRanges.Should().ContainSingle();
    }

    [Theory]
    [InlineData("1Sales")]   // starts with a digit
    [InlineData("My Range")] // contains a space
    [InlineData("A1")]       // collides with a cell reference
    [InlineData("AB12")]     // collides with a cell reference
    [InlineData("net-sales")]// contains a hyphen
    public void AddNamedRange_Rejects_An_Invalid_Name(string badName)
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S");
        var act = () => wb.AddNamedRange(badName, "S!$A$1");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void AddNamedRange_Allows_Underscore_And_Period_And_Digits()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S");
        wb.AddNamedRange("_q1.total", "S!$A$1");
        wb.AddNamedRange("Region2", "S!$B$1");
        wb.NamedRanges.Select(n => n.Name).Should().Equal("_q1.total", "Region2"); // insertion order
    }

    [Fact]
    public void AddNamedRange_Null_And_Empty_Guards()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S");

        ((Action)(() => wb.AddNamedRange(null!, "S!$A$1"))).Should().Throw<ArgumentNullException>();
        ((Action)(() => wb.AddNamedRange("X", null!))).Should().Throw<ArgumentNullException>();
        ((Action)(() => wb.AddNamedRange("", "S!$A$1"))).Should().Throw<ArgumentException>();
        ((Action)(() => wb.AddNamedRange("X", ""))).Should().Throw<ArgumentException>();
    }

    [Fact]
    public void AddNamedRange_Unknown_SheetScope_Throws_SheetNameException()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Real");
        var act = () => wb.AddNamedRange("X", "Real!$A$1", sheetScope: "Ghost");
        act.Should().Throw<SheetNameException>();
    }

    [Fact]
    public void NamedRanges_RoundTrip_Through_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                wb.AddSheet("Data");
                wb.AddSheet("Calc");
                wb.AddNamedRange("Global", "Data!$A$1:$A$100");
                wb.AddNamedRange("Scoped", "Calc!$C$3", sheetScope: "Calc");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                var names = wb.NamedRanges.OrderBy(n => n.Name).ToList();
                names.Should().HaveCount(2);

                names[0].Name.Should().Be("Global");
                names[0].Formula.Should().Be("Data!$A$1:$A$100");
                names[0].SheetScope.Should().BeNull();

                names[1].Name.Should().Be("Scoped");
                names[1].Formula.Should().Be("Calc!$C$3");
                names[1].SheetScope.Should().Be("Calc");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
