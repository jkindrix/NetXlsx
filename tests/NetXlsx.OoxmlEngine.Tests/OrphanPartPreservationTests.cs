// I-82 engine swap — relationship-orphan OPC part preservation
// (pre-cutover parity slice; decision #44 / SDK-quirk #18).
//
// The SDK part graph is relationship-defined: a zip entry with a registered
// content type but NO .rels chain is legal OPC (and exactly what the golden
// RoundTripPreservationTests fixture attaches via raw OPC), but it is
// invisible to SpreadsheetDocument.GetAllParts() and was silently dropped by
// the clone-based Save (O-15 scout finding). The engine now captures orphans
// at Open and re-injects them into every Save. These tests pin that contract
// with fixtures built through System.IO.Packaging (the same no-relationship
// shape as the golden fixture, no NPOI dependency).

using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class OrphanPartPreservationTests
{
    private const string CustomXmlUri = "/customXml/item1.xml";
    private const string PivotCacheUri = "/xl/pivotCache/pivotCacheDefinition1.xml";
    private const string ThreadedCommentsUri = "/xl/threadedComments/threadedComment1.xml";

    private static readonly byte[] CustomXmlPayload = System.Text.Encoding.UTF8.GetBytes(
        "<!-- NetXlsx-preserve-customxml é -->");

    private static readonly byte[] PivotCachePayload = System.Text.Encoding.UTF8.GetBytes(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
        "refreshOnLoad=\"0\" recordCount=\"0\">" +
        "<!-- NetXlsx-preserve-pivotcache -->" +
        "</pivotCacheDefinition>");

    private static readonly byte[] ThreadedCommentsPayload = System.Text.Encoding.UTF8.GetBytes(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\">" +
        "<!-- NetXlsx-preserve-threadedcomments -->" +
        "</ThreadedComments>");

    [Fact]
    public void Orphan_Parts_Survive_Open_Modify_Save_BitIdentical()
    {
        byte[] fixture = BuildFixtureWithOrphans();

        byte[] saved;
        using (var wb = Workbook.OpenOoxml(new MemoryStream(fixture), leaveOpen: false))
        {
            wb["Data"]["A1"].SetString("modified");
            using var outMs = new MemoryStream();
            wb.Save(outMs, leaveOpen: true);
            saved = outMs.ToArray();
        }

        ReadPartBytes(saved, CustomXmlUri).Should().Equal(CustomXmlPayload,
            "design §7.7 / decision #44: an orphan custom-xml part must survive byte-identical");
        ReadPartBytes(saved, PivotCacheUri).Should().Equal(PivotCachePayload,
            "an orphan pivot-cache definition must survive byte-identical");
        ReadPartBytes(saved, ThreadedCommentsUri).Should().Equal(ThreadedCommentsPayload,
            "an orphan threaded-comments part must survive byte-identical");

        // The mutation took effect alongside the preservation.
        using var reopened = Workbook.OpenOoxml(new MemoryStream(saved), leaveOpen: false);
        reopened["Data"]["A1"].GetString().Should().Be("modified");
    }

    [Fact]
    public void Orphan_Parts_Survive_A_Second_Save()
    {
        // Save is repeatable on a live workbook — the orphan snapshot must
        // re-inject on every Save, not just the first.
        byte[] fixture = BuildFixtureWithOrphans();
        using var wb = Workbook.OpenOoxml(new MemoryStream(fixture), leaveOpen: false);

        using (var first = new MemoryStream())
        {
            wb.Save(first, leaveOpen: true);
        }

        wb["Data"]["A1"].SetString("second");
        using var second = new MemoryStream();
        wb.Save(second, leaveOpen: true);
        byte[] saved = second.ToArray();

        ReadPartBytes(saved, CustomXmlUri).Should().Equal(CustomXmlPayload);
        ReadPartBytes(saved, PivotCacheUri).Should().Equal(PivotCachePayload);
        ReadPartBytes(saved, ThreadedCommentsUri).Should().Equal(ThreadedCommentsPayload);
    }

    [Fact]
    public void Orphan_Preserving_Save_Produces_A_Workbook_Both_Read_Paths_Reopen()
    {
        // The injected output must remain a fully valid package — reopen it
        // through the SDK engine and walk its content.
        byte[] fixture = BuildFixtureWithOrphans();
        byte[] saved;
        using (var wb = Workbook.OpenOoxml(new MemoryStream(fixture), leaveOpen: false))
        {
            using var outMs = new MemoryStream();
            wb.Save(outMs, leaveOpen: true);
            saved = outMs.ToArray();
        }

        using var reopened = Workbook.OpenOoxml(new MemoryStream(saved), leaveOpen: false);
        reopened.SheetCount.Should().Be(1);
        reopened["Data"]["A1"].GetString().Should().Be("original");

        // And a second round-trip keeps the orphans (the reopened workbook
        // re-captures them at its own Open).
        using var secondGen = new MemoryStream();
        reopened.Save(secondGen, leaveOpen: true);
        ReadPartBytes(secondGen.ToArray(), CustomXmlUri).Should().Equal(CustomXmlPayload,
            "orphans must survive chained Open->Save->Open->Save round-trips");
    }

    [Fact]
    public void Workbook_Without_Orphans_Is_Unaffected()
    {
        // Regression guard: the no-orphan fast path (created workbooks,
        // rel-wired real-world files) must keep the plain clone Save.
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S")["A1"].SetString("x");
        using var ms = new MemoryStream();
        wb.Save(ms, leaveOpen: true);
        using var reopened = Workbook.OpenOoxml(new MemoryStream(ms.ToArray()), leaveOpen: false);
        reopened["S"]["A1"].GetString().Should().Be("x");
    }

    // ---- Helpers ------------------------------------------------------

    /// <summary>
    /// A valid SDK-engine workbook with three orphan parts attached through
    /// System.IO.Packaging: content types registered, no relationship chain —
    /// the same shape the golden RoundTripPreservationTests fixture builds
    /// via NPOI's raw OPCPackage.
    /// </summary>
    private static byte[] BuildFixtureWithOrphans()
    {
        byte[] baseBytes;
        using (var wb = Workbook.CreateOoxml())
        {
            wb.AddSheet("Data")["A1"].SetString("original");
            using var ms = new MemoryStream();
            wb.Save(ms, leaveOpen: true);
            baseBytes = ms.ToArray();
        }

        var working = new MemoryStream();
        working.Write(baseBytes, 0, baseBytes.Length);
        using (var pkg = Package.Open(working, FileMode.Open, FileAccess.ReadWrite))
        {
            AddOrphan(pkg, CustomXmlUri, "application/xml", CustomXmlPayload);
            AddOrphan(pkg, PivotCacheUri,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
                PivotCachePayload);
            AddOrphan(pkg, ThreadedCommentsUri,
                "application/vnd.ms-excel.threadedcomments+xml",
                ThreadedCommentsPayload);
        }
        return working.ToArray();   // valid on a disposed MemoryStream
    }

    private static void AddOrphan(Package pkg, string uri, string contentType, byte[] payload)
    {
        var part = pkg.CreatePart(new Uri(uri, UriKind.Relative), contentType);
        using var s = part.GetStream(FileMode.Create, FileAccess.Write);
        s.Write(payload, 0, payload.Length);
    }

    private static byte[] ReadPartBytes(byte[] xlsx, string uri)
    {
        using var ms = new MemoryStream(xlsx, writable: false);
        using var pkg = Package.Open(ms, FileMode.Open, FileAccess.Read);
        var partUri = new Uri(uri, UriKind.Relative);
        pkg.PartExists(partUri).Should().BeTrue($"part {uri} must be present in the saved package");
        using var s = pkg.GetPart(partUri).GetStream(FileMode.Open, FileAccess.Read);
        using var buf = new MemoryStream();
        s.CopyTo(buf);
        return buf.ToArray();
    }
}
