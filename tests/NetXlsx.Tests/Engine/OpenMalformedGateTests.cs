// I-82 engine swap — Open malformed-input gate (pre-cutover parity slice).
//
// SDK-quirk #17 (O-15 scout): System.IO.Packaging / SpreadsheetDocument.Open
// happily accept inputs the NPOI engine's I-60 gate rejects — an empty stream
// opens as a brand-new package, a valid zip with no workbook part opens
// without complaint, and package-level corruption surfaces lazily as a raw
// InvalidOperationException at first part access. These tests pin the SDK
// open path's explicit gate so the fuzz harness (NetXlsx.Fuzz) inherits it
// green at cutover.

using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class OpenMalformedGateTests
{
    [Fact]
    public void Empty_Stream_Throws_MalformedFileException()
    {
        // Packaging would otherwise CREATE a new package over the empty
        // read/write stream instead of rejecting it.
        AssertOpenRejects(Array.Empty<byte>());
    }

    [Fact]
    public void Empty_File_Throws_MalformedFileException()
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-empty-{Guid.NewGuid():N}.xlsx");
        File.WriteAllBytes(path, Array.Empty<byte>());
        try
        {
            ((Action)(() => Workbook.Open(path)))
                .Should().Throw<MalformedFileException>();
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Empty_Zip_Throws_MalformedFileException()
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) { }
        AssertOpenRejects(ms.ToArray());
    }

    [Fact]
    public void Zip_With_Random_Entries_Throws_MalformedFileException()
    {
        // A valid zip that is not an OOXML package (no content types, no
        // workbook part).
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            for (int i = 0; i < 3; i++)
            {
                var entry = zip.CreateEntry($"junk{i}.dat");
                using var s = entry.Open();
                s.Write(new byte[] { (byte)i, 0xAA, 0xBB });
            }
        }
        AssertOpenRejects(ms.ToArray());
    }

    [Fact]
    public void Garbage_Bytes_Throw_MalformedFileException()
    {
        var rng = new Random(0xBADF00D);
        var data = new byte[4096];
        rng.NextBytes(data);
        AssertOpenRejects(data);
    }

    [Fact]
    public void Package_Missing_The_Workbook_Part_Wraps_To_MalformedFileException()
    {
        // Deterministic package-level corruption: a structurally valid OOXML
        // zip whose workbook relationship points at a deleted part. The SDK
        // surfaces this lazily as InvalidOperationException ("Specified part
        // does not exist in the package") — the gate must classify it to
        // MalformedFileException, not leak the raw exception.
        byte[] valid = BuildValidWorkbookBytes();
        using var ms = new MemoryStream();
        ms.Write(valid);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("xl/workbook.xml");
            entry.Should().NotBeNull("the fixture must contain the workbook part to delete");
            entry!.Delete();
        }
        AssertOpenRejects(ms.ToArray());
    }

    [Fact]
    public void Zip_With_Truncated_ContentTypes_Throws_MalformedFileException()
    {
        // OOXML-ish (has [Content_Types].xml) but the XML body is malformed.
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = zip.CreateEntry("[Content_Types].xml");
            using var s = entry.Open();
            s.Write(Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><Types><Default"));
        }
        AssertOpenRejects(ms.ToArray());
    }

    [Fact]
    public void Valid_Workbook_Still_Opens_Through_The_Gate()
    {
        // Regression guard: the gate must not over-reject.
        byte[] valid = BuildValidWorkbookBytes();
        using var wb = Workbook.Open(new MemoryStream(valid), leaveOpen: false);
        wb.SheetCount.Should().Be(1);
        wb["S"]["A1"].GetString().Should().Be("hello");
    }

    // ---- Helpers ------------------------------------------------------

    private static void AssertOpenRejects(byte[] data)
    {
        ((Action)(() =>
        {
            using var wb = Workbook.Open(new MemoryStream(data), leaveOpen: false);
        })).Should().Throw<MalformedFileException>();
    }

    private static byte[] BuildValidWorkbookBytes()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].SetString("hello");
        using var ms = new MemoryStream();
        wb.Save(ms, leaveOpen: true);
        return ms.ToArray();
    }
}
