// I-82 engine swap — OPC preservation gate on the SDK engine (lesson #13).
//
// The contract: every OPC part NetXlsx does not model round-trips across an
// Open -> Save on the Open XML SDK engine — nothing silently vanishes. The
// NPOI engine has RoundTripPreservationTests; this is the SDK-engine analogue.
//
// The styles slice observed a 70-85% file-size drop on real OpenOoxml->Save
// round-trips and *assumed* it was harmless deflate recompression. This test
// confirms the assumption structurally: a workbook carrying an unmodeled custom
// XML part (the stand-in for theme/calcChain/printerSettings/customXml/threaded
// comments) preserves that part — present in the output and byte-identical.
//
// The fixture is built programmatically (CI-safe — no dependency on the
// operator's Downloads stress files): create a workbook, attach a custom XML
// part with distinctive bytes via the SDK escape hatch, save. That saved file
// is the "real-world file with an unmodeled part"; the test then opens it on the
// engine, saves it again, and asserts the part set and the custom bytes survive.

using System;
using System.IO;
using System.Linq;
using System.Text;
using AwesomeAssertions;
using DocumentFormat.OpenXml.Packaging;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class OpcPreservationTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-opc-{Guid.NewGuid():N}.xlsx");

    // Distinctive bytes (incl. a multi-byte UTF-8 char) so we know it's *our*
    // payload that round-tripped, not anything the SDK would synthesize.
    private static readonly byte[] CustomPayload =
        Encoding.UTF8.GetBytes("<!-- NetXlsx-opc-preserve é -->");

    private static void BuildFixtureWithCustomPart(string path)
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Data")["A1"].SetString("content");
        var custom = wb.Underlying.WorkbookPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        using (var src = new MemoryStream(CustomPayload))
            custom.FeedData(src);
        wb.Save(path);
    }

    private static string[] PartUris(string path)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        return doc.GetAllParts().Select(p => p.Uri.ToString()).OrderBy(s => s, StringComparer.Ordinal).ToArray();
    }

    private static byte[]? CustomPartBytes(string path)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var cp = doc.GetAllParts().OfType<CustomXmlPart>().FirstOrDefault();
        if (cp is null) return null;
        using var s = cp.GetStream();
        using var mem = new MemoryStream();
        s.CopyTo(mem);
        return mem.ToArray();
    }

    [Fact]
    public void Open_Save_Preserves_The_Full_Part_Set()
    {
        var fixture = TempXlsxPath();
        var roundTripped = TempXlsxPath();
        try
        {
            BuildFixtureWithCustomPart(fixture);
            var before = PartUris(fixture);
            before.Should().Contain(u => u.Contains("customXml"), "the fixture carries an unmodeled custom part");

            using (var wb = Workbook.OpenOoxml(fixture))
                wb.Save(roundTripped);

            var after = PartUris(roundTripped);
            after.Should().Contain(before, "no part NetXlsx does not model may vanish across Open -> Save (lesson #13)");
        }
        finally
        {
            if (File.Exists(fixture)) File.Delete(fixture);
            if (File.Exists(roundTripped)) File.Delete(roundTripped);
        }
    }

    [Fact]
    public void Open_Save_Preserves_Unmodeled_Part_Bytes()
    {
        var fixture = TempXlsxPath();
        var roundTripped = TempXlsxPath();
        try
        {
            BuildFixtureWithCustomPart(fixture);

            using (var wb = Workbook.OpenOoxml(fixture))
                wb.Save(roundTripped);

            CustomPartBytes(roundTripped).Should().Equal(CustomPayload,
                "an unmodeled part round-trips byte-identical, not merely present");
        }
        finally
        {
            if (File.Exists(fixture)) File.Delete(fixture);
            if (File.Exists(roundTripped)) File.Delete(roundTripped);
        }
    }

    [Fact]
    public void Open_Save_Preserves_Modeled_Content_Alongside_Unmodeled_Parts()
    {
        var fixture = TempXlsxPath();
        var roundTripped = TempXlsxPath();
        try
        {
            BuildFixtureWithCustomPart(fixture);

            using (var wb = Workbook.OpenOoxml(fixture))
            {
                // A benign mutation forces a real save path, not a passthrough copy.
                wb["Data"]["B1"].SetNumber(42);
                wb.Save(roundTripped);
            }

            using (var wb = Workbook.OpenOoxml(roundTripped))
            {
                wb["Data"]["A1"].GetString().Should().Be("content");
                wb["Data"]["B1"].GetNumber().Should().Be(42);
            }
            CustomPartBytes(roundTripped).Should().Equal(CustomPayload);
        }
        finally
        {
            if (File.Exists(fixture)) File.Delete(fixture);
            if (File.Exists(roundTripped)) File.Delete(roundTripped);
        }
    }
}
