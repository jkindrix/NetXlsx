// Round-trip preservation test per design §7.7.
//
// **Scope (v0.3.x):** synthetic. Real-world fixtures with pivot caches,
// conditional formatting, threaded comments etc. need to be authored
// in Excel and committed under the fixture-provenance rules
// (decision I18); that's tracked as a v1.0 ship-blocker in the
// roadmap. This test exercises the same contract at lower fidelity:
// it programmatically attaches a custom-XML OPC part to a workbook
// built via raw NPOI, opens the file via NetXlsx, mutates a cell,
// saves, and asserts the custom part survives bit-identical. If
// NetXlsx ever silently drops parts it doesn't model, this test
// catches it.

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NetXlsx;
using FluentAssertions;
using NPOI.OpenXml4Net.OPC;
using NPOI.XSSF.UserModel;
using Xunit;

namespace NetXlsx.GoldenFiles;

public class RoundTripPreservationTests
{
    [Fact]
    public async Task Custom_Opc_Part_Survives_Open_Modify_Save_BitIdentical()
    {
        var path = Path.Combine(Path.GetTempPath(), $"preservation-{Guid.NewGuid():N}.xlsx");
        // A deliberately weird payload so we can be sure it's our custom
        // bytes that round-tripped, not anything NPOI or NetXlsx would
        // synthesize on their own.
        var customPartBytes = new byte[] {
            0x3C, 0x21, 0x2D, 0x2D, // "<!--"
            0x43, 0x4D, 0x41, 0x53, 0x48, 0x45, 0x45, 0x54, 0x53, // "NXLSS"
            0x2D, 0x50, 0x52, 0x45, 0x53, 0x45, 0x52, 0x56, 0x45, // "-PRESERVE"
            0xC3, 0xA9, // "é" (multi-byte UTF-8, to verify binary fidelity)
            0x2D, 0x2D, 0x3E, // "-->"
        };
        var partUriString = "/customXml/item1.xml";

        try
        {
            // ---- Build the input file with a custom OPC part ----------
            using (var wb = new XSSFWorkbook())
            {
                var sheet = wb.CreateSheet("Data");
                sheet.CreateRow(0).CreateCell(0).SetCellValue("original");

                // Attach a custom OPC part directly via the package layer.
                // NetXlsx does not model "custom XML" — that's exactly
                // the kind of part §7.7 promises to preserve.
                var pkg = wb.Package;
                var partUri = PackagingUriHelper.CreatePartName(partUriString);
                var part = pkg.CreatePart(partUri, "application/xml");
                using (var partStream = part.GetOutputStream())
                {
                    partStream.Write(customPartBytes, 0, customPartBytes.Length);
                }

                using var fs = File.Create(path);
                wb.Write(fs, leaveOpen: false);
            }

            // ---- Open via NetXlsx, mutate one cell, save ----------
            using (var wb = await Workbook.OpenAsync(path))
            {
                wb["Data"]["A1"].SetString("modified");
                await wb.SaveAsync(path);
            }

            // ---- Re-open via raw NPOI; assert custom part is intact ---
            using var fs2 = File.OpenRead(path);
            using var afterWb = new XSSFWorkbook(fs2);

            // Cell mutation took effect.
            afterWb.GetSheet("Data").GetRow(0).GetCell(0).StringCellValue.Should().Be("modified");

            // Custom part survived bit-identical.
            var afterPkg = afterWb.Package;
            var partRel = afterPkg.GetParts()
                .FirstOrDefault(p => p.PartName.Name == partUriString);
            partRel.Should().NotBeNull(
                "the custom OPC part must survive Open->modify->Save (design §7.7)");

            byte[] readBack;
            using (var s = partRel!.GetInputStream())
            using (var ms = new MemoryStream())
            {
                s.CopyTo(ms);
                readBack = ms.ToArray();
            }
            readBack.Should().BeEquivalentTo(customPartBytes,
                "the custom part's bytes must be preserved exactly — no transcoding, no normalization");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task NoOp_Open_Save_Does_Not_Mutate_Unmodeled_Parts()
    {
        // Stricter form of the same contract: if the consumer doesn't
        // mutate anything, an Open->Save cycle should not touch unmodeled
        // OPC parts at all. We verify by comparing the custom part's bytes.
        var path = Path.Combine(Path.GetTempPath(), $"preservation-noop-{Guid.NewGuid():N}.xlsx");
        var customBytes = System.Text.Encoding.UTF8.GetBytes(
            "<?xml version=\"1.0\"?><root><untouched>true</untouched></root>");
        var partUriString = "/customXml/item2.xml";

        try
        {
            using (var wb = new XSSFWorkbook())
            {
                wb.CreateSheet("S").CreateRow(0).CreateCell(0).SetCellValue("v");
                var partUri = PackagingUriHelper.CreatePartName(partUriString);
                var part = wb.Package.CreatePart(partUri, "application/xml");
                using var ps = part.GetOutputStream();
                ps.Write(customBytes, 0, customBytes.Length);
                using var fs = File.Create(path);
                wb.Write(fs, leaveOpen: false);
            }

            // Open + immediately save, no mutations.
            using (var wb = await Workbook.OpenAsync(path))
            {
                await wb.SaveAsync(path);
            }

            using var fs2 = File.OpenRead(path);
            using var after = new XSSFWorkbook(fs2);
            var partRel = after.Package.GetParts()
                .FirstOrDefault(p => p.PartName.Name == partUriString);
            partRel.Should().NotBeNull();

            using var s = partRel!.GetInputStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            ms.ToArray().Should().BeEquivalentTo(customBytes);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
