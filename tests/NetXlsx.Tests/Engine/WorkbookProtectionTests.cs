// I-82 engine swap — workbook protection + IsMacroEnabled (formulas/comments/
// hyperlinks slice rider) conformance.
//
// Mirrors the NPOI-engine WorkbookProtectionTests behavioral contract on the
// Open XML SDK engine (decisions I-54 + I-65). Where the NPOI tests reached
// through wb.Underlying (IsStructureLocked etc.), these assert the
// <workbookProtection> DOM via IWorkbook.Underlying — same observable, no
// NPOI dependency. The XOR-verifier hash is asserted byte-for-byte against
// NPOI's CreateXorVerifier1 values ("hunter2" -> C258), pinning cross-engine
// password compatibility.

using System;
using System.IO;
using AwesomeAssertions;
using S = DocumentFormat.OpenXml.Spreadsheet;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class WorkbookProtectionTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-wbprot-{Guid.NewGuid():N}.xlsx");

    private static S.WorkbookProtection? Element(IWorkbook wb)
        => wb.Underlying.WorkbookPart!.Workbook!.GetFirstChild<S.WorkbookProtection>();

    // ---- Behavior -----------------------------------------------------------

    [Fact]
    public void Workbook_Is_Not_Protected_By_Default()
    {
        using var wb = Workbook.Create();
        wb.IsProtected.Should().BeFalse();
        Element(wb).Should().BeNull();
    }

    [Fact]
    public void Protect_Default_Locks_Structure()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Protect();
        wb.IsProtected.Should().BeTrue();

        var wp = Element(wb)!;
        wp.LockStructure!.Value.Should().BeTrue();
        wp.LockWindows!.Value.Should().BeFalse();
        wp.LockRevision!.Value.Should().BeFalse();
    }

    [Fact]
    public void Protect_Applies_All_Three_Flags_When_Specified()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Protect(new WorkbookProtection { Structure = true, Windows = true, Revision = true });

        var wp = Element(wb)!;
        wp.LockStructure!.Value.Should().BeTrue();
        wp.LockWindows!.Value.Should().BeTrue();
        wp.LockRevision!.Value.Should().BeTrue();
    }

    [Fact]
    public void Protect_With_Default_Options_Clears_Unspecified_Flags()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Protect(new WorkbookProtection { Structure = true, Windows = true, Revision = true });
        wb.Protect(WorkbookProtection.LockStructure);

        var wp = Element(wb)!;
        wp.LockStructure!.Value.Should().BeTrue();
        wp.LockWindows!.Value.Should().BeFalse();
        wp.LockRevision!.Value.Should().BeFalse();
    }

    [Fact]
    public void Protect_With_All_False_Options_Reports_Not_Protected()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Protect(WorkbookProtection.Default);
        wb.IsProtected.Should().BeFalse("no flag is locked, matching the NPOI engine");
    }

    [Fact]
    public void Unprotect_Removes_The_Element()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Protect(new WorkbookProtection { Structure = true, Windows = true, Revision = true });
        wb.Unprotect();
        wb.IsProtected.Should().BeFalse();
        Element(wb).Should().BeNull();
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

    // ---- Workbook password (I-65) ---------------------------------------------

    [Fact]
    public void ProtectWithPassword_Sets_The_XOR_Verifier_Hash()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.ProtectWithPassword("hunter2");

        wb.IsProtected.Should().BeTrue();
        // hunter2 -> 0xC258: the same value NPOI's CreateXorVerifier1 writes —
        // cross-engine password compatibility, pinned byte-for-byte.
        Element(wb)!.WorkbookPassword!.Value.Should().Be("C258");
    }

    [Fact]
    public void ProtectWithPassword_Defaults_To_LockStructure()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.ProtectWithPassword("pwd");

        var wp = Element(wb)!;
        wp.LockStructure!.Value.Should().BeTrue();
        wp.LockWindows!.Value.Should().BeFalse();
    }

    [Fact]
    public void ProtectWithPassword_Accepts_Explicit_Options()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.ProtectWithPassword("pwd", new WorkbookProtection { Structure = true, Windows = true });

        var wp = Element(wb)!;
        wp.LockStructure!.Value.Should().BeTrue();
        wp.LockWindows!.Value.Should().BeTrue();
    }

    [Fact]
    public void ProtectWithPassword_Rejects_Null_Password()
    {
        using var wb = Workbook.Create();
        Action act = () => wb.ProtectWithPassword(null!);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Passwordless_Protect_Clears_A_Prior_Password()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.ProtectWithPassword("pwd");
        wb.Protect();   // NPOI parity: ProtectCore(password: null) clears the verifier
        Element(wb)!.WorkbookPassword.Should().BeNull();
        wb.IsProtected.Should().BeTrue();
    }

    // ---- Round-trips -------------------------------------------------------------

    [Fact]
    public void ProtectWithPassword_Survives_File_Roundtrip()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S");
                wb.ProtectWithPassword("hunter2");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb.IsProtected.Should().BeTrue();
                Element(wb)!.WorkbookPassword!.Value.Should().Be("C258");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Protect_Survives_File_Roundtrip()
    {
        var path = TempXlsxPath();
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
                var wp = Element(wb)!;
                wp.LockStructure!.Value.Should().BeTrue();
                wp.LockRevision!.Value.Should().BeTrue();
                wp.LockWindows!.Value.Should().BeFalse();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- IsMacroEnabled ------------------------------------------------------------

    [Fact]
    public void Created_Workbook_Is_Not_MacroEnabled()
    {
        using var wb = Workbook.Create();
        wb.IsMacroEnabled.Should().BeFalse();
    }

    [Fact]
    public void CreateMacroEnabled_Produces_A_MacroEnabled_Workbook_That_Round_Trips()
    {
        // The cutover's conformance pin for the macro-enabled CREATE path
        // (amendment A2): CreateMacroEnabled routes to the SDK engine since
        // v2.0.0, so the created document must carry the macro-enabled
        // document type, report IsMacroEnabled, save with the macro-enabled
        // content type, and round-trip.
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-xlsm-{Guid.NewGuid():N}.xlsm");
        try
        {
            using (var wb = Workbook.CreateMacroEnabled())
            {
                wb.IsMacroEnabled.Should().BeTrue();
                wb.Underlying.DocumentType.Should().Be(
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.MacroEnabledWorkbook);
                wb.AddSheet("S")["A1"].SetString("macro-test");
                wb.Save(path);
            }
            using (var z = System.IO.Compression.ZipFile.OpenRead(path))
            {
                using var r = new StreamReader(z.GetEntry("[Content_Types].xml")!.Open());
                r.ReadToEnd().Should().Contain("macroEnabled.main+xml",
                    "the saved package must carry the macro-enabled workbook content type");
            }
            using (var wb = Workbook.Open(path))
            {
                wb.IsMacroEnabled.Should().BeTrue();
                wb["S"]["A1"].GetString().Should().Be("macro-test");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
