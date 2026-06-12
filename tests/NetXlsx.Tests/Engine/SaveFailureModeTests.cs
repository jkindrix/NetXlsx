// Save failure-mode contract (remediation R-1 / R-2 / R-35).
//
// R-1/R-2: Save(path) is atomic on both engines — serialization targets a
// sibling temp file promoted by rename, so a failed save NEVER truncates or
// destroys a pre-existing destination. (The pre-fix behavior File.Create'd the
// destination before serializing and left a 0-byte husk on failure.)
// R-35: a failed Save must not poison Dispose — with SDK autosave on, Dispose
// re-serialized the invalid DOM, threw the same exception again unhandled at
// the using-brace, and masked the original.
//
// These tests force serialization failure with an XML-illegal control
// character in a cell — today the only in-repo way to fail mid-save. The R-3
// policy slice will change WHERE control characters fail (setter or escaping);
// when it lands, these tests must switch to a different forced failure rather
// than be deleted: the atomicity and dispose contracts they pin are
// failure-source-agnostic.

using System;
using System.IO;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class SaveFailureModeTests
{
    // Pre-I-88 these tests poisoned the save with a control char through
    // SetString; I-88 made that LOSSLESSLY ESCAPED rather than a save-time
    // failure, so the poison now goes through the escape hatch — the one
    // remaining way to get an XML-invalid char into the DOM (and a realistic
    // one: hatch users bypass the facade's encoding). The SDK still throws
    // at serialization, which is exactly the failure these tests pin.
    private static void PoisonCell(ICell cell)
    {
        cell.SetString("poison");
        var text = cell.Underlying.InlineString!
            .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Text>()!;
        text.Text = "poison" + (char)0x01;
    }

    private static string TempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"netxlsx-savefail-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void NoTempResidue(string dir) =>
        Directory.GetFiles(dir, "*.netxlsx.tmp")
            .Should().BeEmpty("temp files must never outlive Save, success or failure");

    // ---- R-35: a failed save must not poison Dispose ------------------------

    [Fact]
    public void FailedSave_DisposeIsClean()
    {
        var wb = Workbook.Create();
        PoisonCell(wb.AddSheet("S")["A1"]);

        using var ms = new MemoryStream();
        Record.Exception(() => wb.Save(ms))
            .Should().NotBeNull("an XML-illegal control character must fail the save");

        Record.Exception(wb.Dispose)
            .Should().BeNull("Dispose after a failed save must not throw (R-35)");
    }

    [Fact]
    public void FailedSave_WorkbookRemainsUsable_AfterRemovingOffendingContent()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        PoisonCell(sheet["A1"]);

        using (var failed = new MemoryStream())
            Record.Exception(() => wb.Save(failed)).Should().NotBeNull();

        sheet["A1"].SetString("recovered");
        using var ok = new MemoryStream();
        wb.Save(ok);
        ok.Position = 0;
        using var reread = Workbook.Open(ok);
        reread[0]["A1"].GetString().Should().Be("recovered",
            "a failed save must not corrupt the live workbook — fixing the content and re-saving must work");
    }

    [Fact]
    public void FailedSave_OnOpenedWorkbook_DisposeIsClean()
    {
        var dir = TempDir();
        try
        {
            var path = Path.Combine(dir, "base.xlsx");
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S")["A1"].SetString("ok");
                wb.Save(path);
            }

            var opened = Workbook.Open(path);
            PoisonCell(opened[0]["A1"]);
            using var ms = new MemoryStream();
            Record.Exception(() => opened.Save(ms)).Should().NotBeNull();
            Record.Exception(opened.Dispose).Should().BeNull();
        }
        finally { Directory.Delete(dir, recursive: true); }
    }

    // ---- R-1: DOM Save(path) atomicity ---------------------------------------

    [Fact]
    public void FailedSavePath_LeavesPreexistingDestinationByteIntact()
    {
        var dir = TempDir();
        try
        {
            var path = Path.Combine(dir, "report.xlsx");
            using (var good = Workbook.Create())
            {
                good.AddSheet("Good")["A1"].SetString("yesterday's report");
                good.Save(path);
            }
            byte[] before = File.ReadAllBytes(path);
            before.Should().NotBeEmpty();

            var bad = Workbook.Create();
            PoisonCell(bad.AddSheet("Bad")["A1"]);
            Record.Exception(() => bad.Save(path)).Should().NotBeNull();
            Record.Exception(bad.Dispose).Should().BeNull();

            File.ReadAllBytes(path).Should().Equal(before,
                "a failed save must leave the destination exactly as it was (R-1)");
            NoTempResidue(dir);
        }
        finally { Directory.Delete(dir, recursive: true); }
    }

    [Fact]
    public void SavePath_OverwritesExistingFile()
    {
        var dir = TempDir();
        try
        {
            var path = Path.Combine(dir, "report.xlsx");
            using (var v1 = Workbook.Create())
            {
                v1.AddSheet("S")["A1"].SetString("version one");
                v1.Save(path);
            }
            using (var v2 = Workbook.Create())
            {
                v2.AddSheet("S")["A1"].SetString("version two");
                v2.Save(path);
            }

            using var reread = Workbook.Open(path);
            reread[0]["A1"].GetString().Should().Be("version two");
            NoTempResidue(dir);
        }
        finally { Directory.Delete(dir, recursive: true); }
    }

    [Fact]
    public void SavePath_MissingDirectory_Throws_AndCreatesNothing()
    {
        var missingDir = Path.Combine(Path.GetTempPath(), $"netxlsx-missing-{Guid.NewGuid():N}");
        var path = Path.Combine(missingDir, "out.xlsx");

        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetString("x");

        Record.Exception(() => wb.Save(path)).Should().BeOfType<DirectoryNotFoundException>();
        Directory.Exists(missingDir).Should().BeFalse("a failed save must not create directories or files");
    }

    // ---- R-2: streaming Save(path) -------------------------------------------

    [Fact]
    public void StreamingSavePath_MissingDirectory_FailsBeforeFinalize_SaveNotBurned()
    {
        using var swb = Workbook.CreateStreaming();
        var sheet = swb.AddSheet("Big");
        sheet.AppendRow().Set(1, "r1");

        var missingDir = Path.Combine(Path.GetTempPath(), $"netxlsx-missing-{Guid.NewGuid():N}");
        Record.Exception(() => swb.Save(Path.Combine(missingDir, "out.xlsx")))
            .Should().BeOfType<DirectoryNotFoundException>();

        // The single-shot save must survive the bad path (the pre-R-2 code's
        // fail-fast guarantee, preserved by the temp name embedding the
        // destination filename).
        var dir = TempDir();
        try
        {
            var path = Path.Combine(dir, "out.xlsx");
            swb.Save(path);
            using var reread = Workbook.Open(path);
            reread[0]["A1"].GetString().Should().Be("r1");
            NoTempResidue(dir);
        }
        finally { Directory.Delete(dir, recursive: true); }
    }

    [Fact]
    public void StreamingSavePath_OverwritesExistingFile()
    {
        var dir = TempDir();
        try
        {
            var path = Path.Combine(dir, "out.xlsx");
            File.WriteAllText(path, "previous contents that must be replaced");

            using var swb = Workbook.CreateStreaming();
            swb.AddSheet("Big").AppendRow().Set(1, "fresh");
            swb.Save(path);

            using var reread = Workbook.Open(path);
            reread[0]["A1"].GetString().Should().Be("fresh");
            NoTempResidue(dir);
        }
        finally { Directory.Delete(dir, recursive: true); }
    }
}
