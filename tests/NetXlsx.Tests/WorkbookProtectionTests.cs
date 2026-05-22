// Coverage for the v1.1 workbook-protection slice: IWorkbook.Protect /
// Unprotect / IsProtected; WorkbookProtection record semantics; file
// round-trip preservation.

using System;
using System.IO;
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
        wb.Underlying.IsStructureLocked().Should().BeTrue();
        wb.Underlying.IsWindowsLocked().Should().BeFalse();
        wb.Underlying.IsRevisionLocked().Should().BeFalse();
    }

    [Fact]
    public void Protect_Applies_All_Three_Flags_When_Specified()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Protect(new WorkbookProtection { Structure = true, Windows = true, Revision = true });
        wb.Underlying.IsStructureLocked().Should().BeTrue();
        wb.Underlying.IsWindowsLocked().Should().BeTrue();
        wb.Underlying.IsRevisionLocked().Should().BeTrue();
    }

    [Fact]
    public void Protect_With_Default_Options_Clears_Unspecified_Flags()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        // First lock everything via raw NPOI, then call Protect with
        // only Structure — windows + revision should clear.
        wb.Underlying.LockWindows();
        wb.Underlying.LockRevision();
        wb.Protect(WorkbookProtection.LockStructure);
        wb.Underlying.IsStructureLocked().Should().BeTrue();
        wb.Underlying.IsWindowsLocked().Should().BeFalse();
        wb.Underlying.IsRevisionLocked().Should().BeFalse();
    }

    [Fact]
    public void Unprotect_Clears_All_Three_Flags()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Protect(new WorkbookProtection { Structure = true, Windows = true, Revision = true });
        wb.Unprotect();
        wb.IsProtected.Should().BeFalse();
        wb.Underlying.IsStructureLocked().Should().BeFalse();
        wb.Underlying.IsWindowsLocked().Should().BeFalse();
        wb.Underlying.IsRevisionLocked().Should().BeFalse();
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
                wb.Underlying.IsStructureLocked().Should().BeTrue();
                wb.Underlying.IsRevisionLocked().Should().BeTrue();
                wb.Underlying.IsWindowsLocked().Should().BeFalse();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
