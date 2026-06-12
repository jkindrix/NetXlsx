// R-14: row/@r and c/@r are OPTIONAL per ECMA-376 §18.3.1.73/§18.3.1.4
// (absent = previous + 1). Excel, LO and openpyxl always emit them, but a
// spec-legal minimal writer may not — and pre-fix, such rows/cells were
// silently invisible to every reader path (data loss on legal input, the
// sharpest I-83 tension in the ledger). The engine now infers and
// materializes missing references once at Open. Cells/rows whose reference
// is PRESENT but corrupt stay untouched (I-83 fail-loud paths own those —
// see MalformedInputContractTests).

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class MissingReferenceNormalizationTests
{
    // Builds a two-row workbook, then splices raw XML into sheet1.xml just
    // before </x:sheetData>.
    private static string FileWithInjectedRows(string rowsXml)
    {
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-r14-{Guid.NewGuid():N}.xlsx");
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("S");
            s["A1"].SetString("one");
            s["A2"].SetString("two");
            wb.Save(path);
        }
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            string xml;
            using (var r = new StreamReader(entry.Open())) xml = r.ReadToEnd();
            xml = xml.Replace("</x:sheetData>", rowsXml + "</x:sheetData>");
            entry.Delete();
            var fresh = zip.CreateEntry("xl/worksheets/sheet1.xml");
            using var w = new StreamWriter(fresh.Open());
            w.Write(xml);
        }
        return path;
    }

    [Fact]
    public void Row_Without_R_Is_Inferred_As_Previous_Plus_One()
    {
        var path = FileWithInjectedRows(
            "<x:row><x:c t=\"inlineStr\"><x:is><x:t>ghost</x:t></x:is></x:c><x:c><x:v>7</x:v></x:c></x:row>");
        try
        {
            using var wb = Workbook.Open(path);
            var s = wb["S"];
            // The @r-less row follows row 2 → inferred row 3; its @r-less
            // cells infer columns A and B.
            s["A3"].GetString().Should().Be("ghost");
            s["B3"].GetNumber().Should().Be(7.0);
            s.LastRowNumber.Should().Be(3, "the inferred row counts toward the extent");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Cell_Without_R_Infers_From_Its_Left_Neighbor()
    {
        var path = FileWithInjectedRows(
            "<x:row r=\"5\"><x:c r=\"B5\"><x:v>1</x:v></x:c><x:c><x:v>2</x:v></x:c></x:row>");
        try
        {
            using var wb = Workbook.Open(path);
            var s = wb["S"];
            s["B5"].GetNumber().Should().Be(1.0);
            s["C5"].GetNumber().Should().Be(2.0, "an @r-less cell is previous + 1 per spec");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Inferred_References_Persist_On_Resave_And_Feed_Dimension()
    {
        var path = FileWithInjectedRows(
            "<x:row><x:c><x:v>9</x:v></x:c></x:row>");
        try
        {
            using var ms = new MemoryStream();
            using (var wb = Workbook.Open(path))
            {
                wb.Save(ms);
            }
            ms.Position = 0;
            using (var wb = Workbook.Open(ms))
            {
                wb["S"]["A3"].GetNumber().Should().Be(9.0);
                SavedOoxml.SheetXml(wb).Root!
                    .Element(SavedOoxml.Main + "dimension")!.Attribute("ref")!.Value
                    .Should().Be("A1:A3", "the normalized grid feeds the R-13 dimension");
            }
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Corrupt_Cell_Reference_Stops_Inference_And_Fails_Loud_On_Access()
    {
        // A PRESENT-but-unparseable @r is corruption, not omission —
        // normalization leaves it untouched and stops inferring for the
        // rest of the row (the running column is no longer trustworthy).
        // Accessing the row then fails loud through the existing I-83
        // path rather than silently substituting anything.
        var path = FileWithInjectedRows(
            "<x:row r=\"5\"><x:c r=\"NOT-A-REF\"><x:v>1</x:v></x:c><x:c><x:v>2</x:v></x:c></x:row>");
        try
        {
            using var wb = Workbook.Open(path);
            var s = wb["S"];
            Action act = () => _ = s["A5"].Kind;
            act.Should().Throw<InvalidCellAddressException>(
                "a corrupt cell reference is file corruption — fail loud, never infer over it");
        }
        finally { File.Delete(path); }
    }
}
