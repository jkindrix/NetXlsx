// I-82 engine swap — structure slice (5b): sheet-protection conformance.
// Mirrors the NPOI engine's ISheet.Protect / Unprotect / IsProtected contract
// (decision I-53) on the Open XML SDK engine, reading the <sheetProtection> node
// directly: @sheet flags protection, the 15 granular lock attributes reflect the
// SheetProtection options, and @password carries the legacy 16-bit XOR verifier.
// Cross-checked against NetXlsx.Tests.SheetProtectionTests.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.OoxmlEngine.Tests;

public class SheetProtectionTests
{
    private static S.SheetProtection? Prot(IWorkbook wb, string sheet)
    {
        var part = wb.Underlying.WorkbookPart!.WorksheetParts;
        // Single-sheet fixtures use the only worksheet part; multi-sheet resolve by
        // matching the worksheet that owns a <sheetProtection>. Tests below use one
        // protected sheet at a time, so the first part with protection is correct.
        foreach (var p in part)
        {
            var sp = p.Worksheet!.GetFirstChild<S.SheetProtection>();
            if (sp is not null) return sp;
        }
        return null;
    }

    [Fact]
    public void Sheet_Is_Not_Protected_By_Default()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S").IsProtected.Should().BeFalse();
    }

    [Fact]
    public void Protect_Without_Password_Flags_The_Sheet()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.Protect();
        s.IsProtected.Should().BeTrue();
        Prot(wb, "S")!.Sheet!.Value.Should().BeTrue();
        Prot(wb, "S")!.Password.Should().BeNull("no password was supplied");
    }

    [Fact]
    public void Protect_With_Password_Writes_A_Verifier()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.Protect(password: "hunter2");
        s.IsProtected.Should().BeTrue();
        Prot(wb, "S")!.Password!.Value.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void Unprotect_Removes_The_Element()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.Protect();
        s.IsProtected.Should().BeTrue();
        s.Unprotect();
        s.IsProtected.Should().BeFalse();
        Prot(wb, "S").Should().BeNull("Unprotect removes the <sheetProtection> element");
    }

    [Fact]
    public void Unprotect_On_Unprotected_Sheet_Is_Safe()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        ((Action)(() => s.Unprotect())).Should().NotThrow();
        s.IsProtected.Should().BeFalse();
    }

    [Fact]
    public void Protect_Applies_Granular_Lock_Flags()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.Protect(options: SheetProtection.LockAll);

        var sp = Prot(wb, "S")!;
        sp.FormatCells!.Value.Should().BeTrue();
        sp.FormatColumns!.Value.Should().BeTrue();
        sp.InsertRows!.Value.Should().BeTrue();
        sp.DeleteColumns!.Value.Should().BeTrue();
        sp.Sort!.Value.Should().BeTrue();
        sp.AutoFilter!.Value.Should().BeTrue();
        sp.Objects!.Value.Should().BeTrue();
        sp.Scenarios!.Value.Should().BeTrue();
    }

    [Fact]
    public void Protect_With_Partial_Options_Sets_Only_Specified_Flags()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.Protect(options: new SheetProtection { LockSort = true });

        var sp = Prot(wb, "S")!;
        sp.Sort!.Value.Should().BeTrue();
        sp.FormatCells!.Value.Should().BeFalse();
        sp.AutoFilter!.Value.Should().BeFalse();
    }

    [Fact]
    public void Protect_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-prot-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var s = wb.AddSheet("S");
                s["A1"].SetString("content");
                s.Protect(options: new SheetProtection { LockFormatCells = true, LockSort = true });
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                var s = wb["S"];
                s.IsProtected.Should().BeTrue();
                var sp = Prot(wb, "S")!;
                sp.FormatCells!.Value.Should().BeTrue();
                sp.Sort!.Value.Should().BeTrue();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Protect_Lands_Before_MergeCells_In_Schema_Order()
    {
        // <sheetProtection> sits before <mergeCells> in CT_Worksheet — adding a
        // merge first then protecting must still produce valid ordering.
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s.MergeCells("A1:C1");
        s.Protect(password: "pw");

        var ws = wb.Underlying.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        var children = ws.ChildElements.ToList();
        int spIdx = children.FindIndex(e => e is S.SheetProtection);
        int mcIdx = children.FindIndex(e => e is S.MergeCells);
        spIdx.Should().BeGreaterThan(-1);
        mcIdx.Should().BeGreaterThan(-1);
        spIdx.Should().BeLessThan(mcIdx, "<sheetProtection> must precede <mergeCells>");
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Protection_Is_Schema_Valid()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("A").Protect();
        wb.AddSheet("B").Protect(password: "secret", options: SheetProtection.LockAll);
        OpenXmlValidationGate.AssertValid(wb);
    }
}
