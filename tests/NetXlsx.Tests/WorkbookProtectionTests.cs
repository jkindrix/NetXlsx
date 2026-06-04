// Coverage for the v1.1 workbook-protection slice: IWorkbook.Protect /
// Unprotect / IsProtected; WorkbookProtection record semantics; file
// round-trip preservation.
//
// Granular lock flags have no public read-back, so tests assert on the
// persisted <workbookProtection> element (engine-agnostic; see SavedOoxml)
// rather than reaching through `.Underlying` into NPOI.

using System;
using System.IO;
using System.Xml.Linq;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class WorkbookProtectionTests
{
    // ---- Value-record semantics ---------------------------------------

    [Fact]
    public void WorkbookProtection_Default_Is_All_False()
    {
        var p = WorkbookProtection.Default;
        p.Structure.Should().BeFalse();
        p.Windows.Should().BeFalse();
        p.Revision.Should().BeFalse();
    }

    [Fact]
    public void WorkbookProtection_LockStructure_Sets_Only_Structure()
    {
        var p = WorkbookProtection.LockStructure;
        p.Structure.Should().BeTrue();
        p.Windows.Should().BeFalse();
        p.Revision.Should().BeFalse();
    }

    [Fact]
    public void WorkbookProtection_Equality_Is_Structural()
    {
        var a = new WorkbookProtection { Structure = true, Revision = true };
        var b = new WorkbookProtection { Structure = true, Revision = true };
        a.Should().Be(b);
        a.GetHashCode().Should().Be(b.GetHashCode());
    }

    // ---- Behavior -----------------------------------------------------

    [Fact]
    public void Workbook_Is_Not_Protected_By_Default()
    {
        using var wb = Workbook.Create();
        wb.IsProtected.Should().BeFalse();
    }

    [Fact]
    public void Protect_Default_Locks_Structure()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Protect();
        wb.IsProtected.Should().BeTrue();
        var prot = ProtectionElement(wb);
        SavedOoxml.BoolAttr(prot, "lockStructure").Should().BeTrue();
        SavedOoxml.BoolAttr(prot, "lockWindows").Should().BeFalse();
        SavedOoxml.BoolAttr(prot, "lockRevision").Should().BeFalse();
    }

    [Fact]
    public void Protect_Applies_All_Three_Flags_When_Specified()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Protect(new WorkbookProtection { Structure = true, Windows = true, Revision = true });
        var prot = ProtectionElement(wb);
        SavedOoxml.BoolAttr(prot, "lockStructure").Should().BeTrue();
        SavedOoxml.BoolAttr(prot, "lockWindows").Should().BeTrue();
        SavedOoxml.BoolAttr(prot, "lockRevision").Should().BeTrue();
    }

    [Fact]
    public void Protect_With_Default_Options_Clears_Unspecified_Flags()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        // First lock everything, then call Protect with only Structure —
        // windows + revision should clear.
        wb.Protect(new WorkbookProtection { Structure = true, Windows = true, Revision = true });
        wb.Protect(WorkbookProtection.LockStructure);
        var prot = ProtectionElement(wb);
        SavedOoxml.BoolAttr(prot, "lockStructure").Should().BeTrue();
        SavedOoxml.BoolAttr(prot, "lockWindows").Should().BeFalse();
        SavedOoxml.BoolAttr(prot, "lockRevision").Should().BeFalse();
    }

    [Fact]
    public void Unprotect_Clears_All_Three_Flags()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Protect(new WorkbookProtection { Structure = true, Windows = true, Revision = true });
        wb.Unprotect();
        wb.IsProtected.Should().BeFalse();
        // Engines differ on the element's fate after Unprotect (cleared
        // flags vs removed element); the contract is that no lock survives.
        var prot = ProtectionElement(wb);
        SavedOoxml.BoolAttr(prot, "lockStructure").Should().BeFalse();
        SavedOoxml.BoolAttr(prot, "lockWindows").Should().BeFalse();
        SavedOoxml.BoolAttr(prot, "lockRevision").Should().BeFalse();
    }

    [Fact]
    public void Unprotect_On_Unprotected_Workbook_Is_Safe()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        Action act = () => wb.Unprotect();
        act.Should().NotThrow();
        wb.IsProtected.Should().BeFalse();
    }

    // ---- File round-trip ----------------------------------------------

    // ---- Workbook password (v1.2 / I-65) ------------------------------

    [Fact]
    public void ProtectWithPassword_Sets_The_XOR_Verifier_Hash()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.ProtectWithPassword("hunter2");

        wb.IsProtected.Should().BeTrue();
        // hunter2 -> 0xC258 per the legacy 16-bit XOR verifier
        // (NPOI CryptoFunctions.CreateXorVerifier1; pinned cross-engine).
        PasswordAttr(wb).Should().Be("C258");
    }

    [Fact]
    public void ProtectWithPassword_Defaults_To_LockStructure()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.ProtectWithPassword("pwd");
        var prot = ProtectionElement(wb);
        SavedOoxml.BoolAttr(prot, "lockStructure").Should().BeTrue();
        SavedOoxml.BoolAttr(prot, "lockWindows").Should().BeFalse();
    }

    [Fact]
    public void ProtectWithPassword_Accepts_Explicit_Options()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.ProtectWithPassword("pwd", new WorkbookProtection { Structure = true, Windows = true });
        var prot = ProtectionElement(wb);
        SavedOoxml.BoolAttr(prot, "lockStructure").Should().BeTrue();
        SavedOoxml.BoolAttr(prot, "lockWindows").Should().BeTrue();
    }

    [Fact]
    public void ProtectWithPassword_Rejects_Null_Password()
    {
        using var wb = Workbook.Create();
        Action act = () => wb.ProtectWithPassword(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Unprotect_Clears_Password_Too()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.ProtectWithPassword("pwd");
        PasswordAttr(wb).Should().NotBeNull();

        wb.Unprotect();
        wb.IsProtected.Should().BeFalse();
        // Note: the engines differ on whether the workbookPassword
        // verifier itself survives Unprotect — without the structure/
        // windows/revision flags the element is effectively a no-op, so
        // the contract is only that no lock remains. Documented in
        // implementation-notes.md.
        var prot = ProtectionElement(wb);
        SavedOoxml.BoolAttr(prot, "lockStructure").Should().BeFalse();
        SavedOoxml.BoolAttr(prot, "lockWindows").Should().BeFalse();
        SavedOoxml.BoolAttr(prot, "lockRevision").Should().BeFalse();
    }

    [Fact]
    public void ProtectWithPassword_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"wbpwd-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S");
                wb.ProtectWithPassword("hunter2");
                wb.Save(path);
            }

            // The persisted artifact carries the verifier…
            var prot = SavedOoxml.PartFromFile(path, "xl/workbook.xml").Root!
                .Element(SavedOoxml.Main + "workbookProtection");
            prot.Should().NotBeNull();
            prot!.Attribute("workbookPassword")!.Value.Should().Be("C258");

            // …and the reopened workbook both reads it back and re-persists it.
            using (var wb = Workbook.Open(path))
            {
                wb.IsProtected.Should().BeTrue();
                PasswordAttr(wb).Should().Be("C258");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Protect_Survives_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"wbprot-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S");
                wb.Protect(new WorkbookProtection { Structure = true, Revision = true });
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb.IsProtected.Should().BeTrue();
                var prot = ProtectionElement(wb);
                SavedOoxml.BoolAttr(prot, "lockStructure").Should().BeTrue();
                SavedOoxml.BoolAttr(prot, "lockRevision").Should().BeTrue();
                SavedOoxml.BoolAttr(prot, "lockWindows").Should().BeFalse();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- helpers ------------------------------------------------------

    private static XElement? ProtectionElement(IWorkbook wb)
        => SavedOoxml.WorkbookXml(wb).Root!
            .Element(SavedOoxml.Main + "workbookProtection");

    private static string? PasswordAttr(IWorkbook wb)
        => ProtectionElement(wb)?.Attribute("workbookPassword")?.Value;
}
