// Round-trip preservation test per design §7.7 / decision #44.
//
// The contract: every OPC part NetXlsx does not model is preserved
// byte-identical across Open -> Modify -> Save. This is the single
// highest-value invariant the library promises beyond "produces a
// valid xlsx."
//
// Four categories from decision #44 are exercised here:
//
//   1. Custom XML        (/customXml/item*.xml)
//   2. Conditional formatting (embedded in worksheet XML)
//   3. Pivot caches      (/xl/pivotCache/pivotCacheDefinition*.xml + records)
//   4. Threaded comments (/xl/threadedComments/threadedComment*.xml)
//
// The fixture is built programmatically (per decision I18 option b —
// "script-generated, each has a sibling .gen.cs that produces it on
// demand"). BuildFixture() below is that script. Categories 1, 3, 4
// are attached via raw OPC because they're either unmodeled by NPOI's
// high-level API (threaded comments — Excel 365 feature) or too narrow
// to author via NPOI's API (pivot caches). Category 2 uses NPOI's
// high-level conditional-formatting API so it serializes into the
// worksheet XML the way Excel itself would.
//
// The XML payloads attached for categories 1, 3, 4 are intentionally
// minimal (well-formed but semantically inert). The test verifies
// preservation of bytes, not of Excel-pivot-table behavior.

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NetXlsx;
using AwesomeAssertions;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace NetXlsx.GoldenFiles;

public class RoundTripPreservationTests
{
    // ------------------------------------------------------------------
    // Fixture builder + helpers
    // ------------------------------------------------------------------

    private const string CustomXmlPartUri = "/customXml/item1.xml";
    private const string PivotCacheDefnPartUri = "/xl/pivotCache/pivotCacheDefinition1.xml";
    private const string ThreadedCommentsPartUri = "/xl/threadedComments/threadedComment1.xml";

    private static readonly byte[] CustomXmlPayload = System.Text.Encoding.UTF8.GetBytes(
        // Marked with deliberate distinguishable bytes, including a multi-byte
        // UTF-8 character, so we can be sure it's our bytes that round-tripped
        // (not anything NPOI or NetXlsx would synthesize on their own).
        "<!-- NetXlsx-preserve-customxml é -->");

    private static readonly byte[] PivotCacheDefnPayload = System.Text.Encoding.UTF8.GetBytes(
        // Minimal pivot-cache-definition stub: well-formed XML with the
        // expected namespace, intentionally empty.
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
        "refreshOnLoad=\"0\" recordCount=\"0\">" +
        "<!-- NetXlsx-preserve-pivotcache -->" +
        "</pivotCacheDefinition>");

    private static readonly byte[] ThreadedCommentsPayload = System.Text.Encoding.UTF8.GetBytes(
        // Minimal threaded-comments stub. The Excel 365 namespace is
        // x:thc; the structure is well-formed but intentionally empty.
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\">" +
        "<!-- NetXlsx-preserve-threadedcomments -->" +
        "</ThreadedComments>");

    /// <summary>
    /// Builds the preservation fixture: a workbook with one regular sheet
    /// plus all four unmodeled-part categories attached. The caller owns
    /// <paramref name="path"/> cleanup.
    /// </summary>
    private static void BuildFixture(string path)
    {
        using var wb = new XSSFWorkbook();
        var sheet = wb.CreateSheet("Data");
        for (int r = 0; r < 5; r++)
        {
            var row = sheet.CreateRow(r);
            row.CreateCell(0).SetCellValue($"row {r}");
            row.CreateCell(1).SetCellValue((r + 1) * 25.0);
        }

        // ---- Category 2: conditional formatting (high-level NPOI API) ----
        // A simple "highlight if greater than 50" rule on B1:B5. NPOI
        // serializes this into the worksheet XML.
        var scf = sheet.SheetConditionalFormatting;
        var rule = scf.CreateConditionalFormattingRule(
            ComparisonOperator.GreaterThan, "50");
        var fontFmt = rule.CreateFontFormatting();
        fontFmt.SetFontStyle(italic: true, bold: true);
        scf.AddConditionalFormatting(
            new[] { CellRangeAddress.ValueOf("B1:B5") },
            rule);

        // ---- Categories 1, 3, 4: raw OPC attachments ----
        var pkg = wb.Package;
        AddOpcPart(pkg, CustomXmlPartUri, "application/xml", CustomXmlPayload);
        AddOpcPart(pkg, PivotCacheDefnPartUri,
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
                   PivotCacheDefnPayload);
        AddOpcPart(pkg, ThreadedCommentsPartUri,
                   "application/vnd.ms-excel.threadedcomments+xml",
                   ThreadedCommentsPayload);

        using var fs = File.Create(path);
        wb.Write(fs, leaveOpen: false);
    }

    private static void AddOpcPart(OPCPackage pkg, string uri, string contentType, byte[] payload)
    {
        var partName = PackagingUriHelper.CreatePartName(uri);
        var part = pkg.CreatePart(partName, contentType);
        using var s = part.GetOutputStream();
        s.Write(payload, 0, payload.Length);
    }

    private static byte[] ReadPartBytes(string xlsxPath, string partUri)
    {
        var pkg = OPCPackage.Open(xlsxPath, PackageAccess.READ);
        try
        {
            var name = PackagingUriHelper.CreatePartName(partUri);
            if (!pkg.ContainPart(name))
                throw new InvalidOperationException($"part not found: {partUri}");
            var part = pkg.GetPart(name);
            using var s = part.GetInputStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            return ms.ToArray();
        }
        finally
        {
            pkg.Close();
        }
    }

    private static string ReadSheetXml(string xlsxPath, string sheetPartUri = "/xl/worksheets/sheet1.xml")
    {
        var bytes = ReadPartBytes(xlsxPath, sheetPartUri);
        return System.Text.Encoding.UTF8.GetString(bytes);
    }

    // ------------------------------------------------------------------
    // The contract: all four part types survive Open -> Modify -> Save
    // ------------------------------------------------------------------

    [Fact]
    public async Task All_Four_Unmodeled_Part_Types_Survive_Open_Modify_Save()
    {
        var path = Path.Combine(Path.GetTempPath(), $"preservation-all-{Guid.NewGuid():N}.xlsx");
        try
        {
            BuildFixture(path);

            // Sanity-check the fixture: all four are present pre-modify.
            ReadPartBytes(path, CustomXmlPartUri).Should().BeEquivalentTo(CustomXmlPayload,
                "fixture sanity: custom-xml part must exist before Open->Save");
            ReadPartBytes(path, PivotCacheDefnPartUri).Should().BeEquivalentTo(PivotCacheDefnPayload,
                "fixture sanity: pivot-cache part must exist before Open->Save");
            ReadPartBytes(path, ThreadedCommentsPartUri).Should().BeEquivalentTo(ThreadedCommentsPayload,
                "fixture sanity: threaded-comments part must exist before Open->Save");
            ReadSheetXml(path).Should().Contain("conditionalFormatting",
                "fixture sanity: CF rule must be in worksheet XML before Open->Save");

            // ---- Open through NetXlsx, mutate one cell, save ----
            using (var wb = await Workbook.OpenAsync(path))
            {
                wb["Data"]["A1"].SetString("modified");
                await wb.SaveAsync(path);
            }

            // ---- All four part types should still be present and
            // byte-identical (except the worksheet sheet1.xml, which we
            // mutated; CF check there is by-substring, not by-bytes) ----
            ReadPartBytes(path, CustomXmlPartUri).Should().BeEquivalentTo(CustomXmlPayload,
                "design §7.7: custom-xml part must survive Open->Modify->Save byte-identical");
            ReadPartBytes(path, PivotCacheDefnPartUri).Should().BeEquivalentTo(PivotCacheDefnPayload,
                "design §7.7: pivot-cache definition part must survive byte-identical");
            ReadPartBytes(path, ThreadedCommentsPartUri).Should().BeEquivalentTo(ThreadedCommentsPayload,
                "design §7.7: threaded-comments part must survive byte-identical");

            // Conditional formatting lives inside the worksheet XML, which
            // we DID mutate (cell A1 changed). Verify the CF rule survived
            // by checking the worksheet XML still contains the structural
            // markers and our distinguishing "50" threshold value.
            var sheetXml = ReadSheetXml(path);
            sheetXml.Should().Contain("conditionalFormatting",
                "design §7.7: CF rule must survive worksheet round-trip");
            sheetXml.Should().Contain(">50<",
                "the specific CF threshold ('greater than 50') must survive — not just an empty CF block");

            // The actual cell mutation also took effect.
            using (var fs = File.OpenRead(path))
            using (var ws = new XSSFWorkbook(fs))
            {
                ws.GetSheet("Data").GetRow(0).GetCell(0).StringCellValue
                    .Should().Be("modified");
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task Noop_Open_Save_Does_Not_Mutate_Any_Of_The_Four_Part_Types()
    {
        // Stricter form of the same contract: if the consumer doesn't
        // mutate anything, an Open->Save should not touch unmodeled parts
        // at all. CF lives in the worksheet XML which IS rewritten on
        // every Save (NPOI re-serializes), so we don't byte-check it
        // here — we still substring-check that CF survives.
        var path = Path.Combine(Path.GetTempPath(), $"preservation-noop-{Guid.NewGuid():N}.xlsx");
        try
        {
            BuildFixture(path);

            using (var wb = await Workbook.OpenAsync(path))
            {
                await wb.SaveAsync(path);
            }

            ReadPartBytes(path, CustomXmlPartUri).Should().BeEquivalentTo(CustomXmlPayload);
            ReadPartBytes(path, PivotCacheDefnPartUri).Should().BeEquivalentTo(PivotCacheDefnPayload);
            ReadPartBytes(path, ThreadedCommentsPartUri).Should().BeEquivalentTo(ThreadedCommentsPayload);
            ReadSheetXml(path).Should().Contain("conditionalFormatting");
            ReadSheetXml(path).Should().Contain(">50<");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ------------------------------------------------------------------
    // Single-part backwards compatibility — preserves the original v0.3.x
    // narrow test as a smoke check that the simpler-case still works.
    // ------------------------------------------------------------------

    [Fact]
    public async Task Single_Custom_Opc_Part_Survives_Open_Modify_Save_BitIdentical()
    {
        var path = Path.Combine(Path.GetTempPath(), $"preservation-single-{Guid.NewGuid():N}.xlsx");
        var customPartBytes = new byte[] {
            0x3C, 0x21, 0x2D, 0x2D,
            0x4E, 0x45, 0x54, 0x58, 0x4C, 0x53, 0x58,           // "NETXLSX"
            0x2D, 0x50, 0x52, 0x45, 0x53, 0x45, 0x52, 0x56, 0x45,
            0xC3, 0xA9,
            0x2D, 0x2D, 0x3E,
        };
        var partUri = "/customXml/itemSingle.xml";

        try
        {
            using (var wb = new XSSFWorkbook())
            {
                var sheet = wb.CreateSheet("Data");
                sheet.CreateRow(0).CreateCell(0).SetCellValue("original");
                AddOpcPart(wb.Package, partUri, "application/xml", customPartBytes);
                using var fs = File.Create(path);
                wb.Write(fs, leaveOpen: false);
            }

            using (var wb = await Workbook.OpenAsync(path))
            {
                wb["Data"]["A1"].SetString("modified");
                await wb.SaveAsync(path);
            }

            ReadPartBytes(path, partUri).Should().BeEquivalentTo(customPartBytes);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
