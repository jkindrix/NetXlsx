// Coverage for the v1.1 sheet-protection slice: ISheet.Protect /
// Unprotect / IsProtected; SheetProtection record semantics; file
// round-trip preservation; granular lock flags propagated to the
// persisted <sheetProtection> element.
//
// OOXML's lock attributes have INVERTED defaults: the format/insert/
// delete/sort/autoFilter/pivotTables family defaults to TRUE (locked)
// when absent, so "locked" often serializes as an OMITTED attribute and
// "unlocked" as an explicit "0". LockFlag encodes that.

using System;
using System.IO;
using System.Xml.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class SheetProtectionTests
{
    // ---- Value-record semantics ---------------------------------------

    [Fact]
    public void SheetProtection_Default_Is_All_False()
    {
        var p = SheetProtection.Default;
        p.LockFormatCells.Should().BeFalse();
        p.LockFormatColumns.Should().BeFalse();
        p.LockInsertRows.Should().BeFalse();
        p.LockDeleteColumns.Should().BeFalse();
        p.LockSelectLockedCells.Should().BeFalse();
        p.LockSort.Should().BeFalse();
        p.LockAutoFilter.Should().BeFalse();
        p.LockPivotTables.Should().BeFalse();
        p.LockObjects.Should().BeFalse();
        p.LockScenarios.Should().BeFalse();
    }

    [Fact]
    public void SheetProtection_LockAll_Is_All_True()
    {
        var p = SheetProtection.LockAll;
        p.LockFormatCells.Should().BeTrue();
        p.LockFormatColumns.Should().BeTrue();
        p.LockFormatRows.Should().BeTrue();
        p.LockInsertColumns.Should().BeTrue();
        p.LockInsertRows.Should().BeTrue();
        p.LockInsertHyperlinks.Should().BeTrue();
        p.LockDeleteColumns.Should().BeTrue();
        p.LockDeleteRows.Should().BeTrue();
        p.LockSelectLockedCells.Should().BeTrue();
        p.LockSelectUnlockedCells.Should().BeTrue();
        p.LockSort.Should().BeTrue();
        p.LockAutoFilter.Should().BeTrue();
        p.LockPivotTables.Should().BeTrue();
        p.LockObjects.Should().BeTrue();
        p.LockScenarios.Should().BeTrue();
    }

    [Fact]
    public void SheetProtection_Equality_Is_Structural()
    {
        var a = new SheetProtection { LockFormatCells = true, LockSort = true };
        var b = new SheetProtection { LockFormatCells = true, LockSort = true };
        a.Should().Be(b);
        a.GetHashCode().Should().Be(b.GetHashCode());
    }

    // ---- Behavior -----------------------------------------------------

    [Fact]
    public void Sheet_Is_Not_Protected_By_Default()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.IsProtected.Should().BeFalse();
    }

    [Fact]
    public void Protect_Without_Password_Flags_The_Sheet()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.Protect();
        sh.IsProtected.Should().BeTrue();
    }

    [Fact]
    public void Protect_With_Password_Flags_The_Sheet()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.Protect(password: "hunter2");
        sh.IsProtected.Should().BeTrue();
    }

    [Fact]
    public void Unprotect_Clears_The_Flag()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.Protect();
        sh.IsProtected.Should().BeTrue();
        sh.Unprotect();
        sh.IsProtected.Should().BeFalse();
    }

    [Fact]
    public void Unprotect_On_Unprotected_Sheet_Is_Safe()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.Unprotect();
        act.Should().NotThrow();
        sh.IsProtected.Should().BeFalse();
    }

    [Fact]
    public void Protect_Applies_Granular_Lock_Flags()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.Protect(options: SheetProtection.LockAll);

        // Inspect the persisted sheet-protection element to confirm the
        // flags were propagated.
        var sp = ProtectionElement(wb);
        sp.Should().NotBeNull();
        LockFlag(sp!, "formatCells").Should().BeTrue();
        LockFlag(sp!, "formatColumns").Should().BeTrue();
        LockFlag(sp!, "insertRows").Should().BeTrue();
        LockFlag(sp!, "deleteColumns").Should().BeTrue();
        LockFlag(sp!, "sort").Should().BeTrue();
        LockFlag(sp!, "autoFilter").Should().BeTrue();
    }

    [Fact]
    public void Protect_With_Partial_Options_Sets_Only_Specified_Flags()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh.Protect(options: new SheetProtection { LockSort = true });

        var sp = ProtectionElement(wb);
        sp.Should().NotBeNull();
        LockFlag(sp!, "sort").Should().BeTrue();
        // Other flags remain unlocked.
        LockFlag(sp!, "formatCells").Should().BeFalse();
        LockFlag(sp!, "autoFilter").Should().BeFalse();
    }

    // ---- File round-trip ----------------------------------------------

    [Fact]
    public void Protect_Survives_File_Roundtrip_Without_Password()
    {
        var path = Path.Combine(Path.GetTempPath(), $"prot-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh["A1"].SetString("content");
                sh.Protect(options: new SheetProtection
                {
                    LockFormatCells = true,
                    LockSort = true,
                });
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var sh = wb["S"];
                sh.IsProtected.Should().BeTrue();
                var sp = ProtectionElement(wb);
                sp.Should().NotBeNull();
                LockFlag(sp!, "formatCells").Should().BeTrue();
                LockFlag(sp!, "sort").Should().BeTrue();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Protect_Survives_File_Roundtrip_With_Password()
    {
        var path = Path.Combine(Path.GetTempPath(), $"prot-pwd-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh["A1"].SetString("content");
                sh.Protect(password: "hunter2");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb["S"].IsProtected.Should().BeTrue();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- helpers ------------------------------------------------------

    private static XElement? ProtectionElement(IWorkbook wb)
        => SavedOoxml.SheetXml(wb).Root!
            .Element(SavedOoxml.Main + "sheetProtection");

    /// <summary>
    /// Reads a sheetProtection lock attribute honoring its OOXML schema
    /// default: the format/insert/delete/sort/autoFilter/pivotTables
    /// family defaults to TRUE (locked) when the attribute is absent.
    /// </summary>
    private static bool LockFlag(XElement sp, string attribute, bool oxmlDefault = true)
        => sp.Attribute(attribute) is { } a
            ? a.Value is "1" or "true"
            : oxmlDefault;

}
