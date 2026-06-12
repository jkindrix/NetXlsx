// R-13: both writers emit <dimension> (the declared used range) so sized
// consumers work — openpyxl's read-only/streaming mode refused the pre-fix
// streaming output outright ("Worksheet is unsized"). The DOM engine
// refreshes the element from the live extent on every Save (also correcting
// a stale dimension carried by an opened file); the streaming engine tracks
// the extent as rows flush and splices the element in at assembly, since
// the CT_Worksheet sequence puts dimension BEFORE sheetData but the value
// is only known after the last row.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class DimensionTests
{
    private static string? DimensionRef(IWorkbook wb, int sheetNumber = 1)
        => SavedOoxml.SheetXml(wb, sheetNumber).Root!
            .Element(SavedOoxml.Main + "dimension")?.Attribute("ref")?.Value;

    [Fact]
    public void Dom_Save_Emits_Dimension_Covering_The_Used_Range()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["B2"].SetString("x");
        s["D7"].SetNumber(1);

        DimensionRef(wb).Should().Be("B2:D7");
    }

    [Fact]
    public void Dom_Save_Emits_A1_For_An_Empty_Sheet()
    {
        // Excel's own convention for a sheet with no cells.
        using var wb = Workbook.Create();
        wb.AddSheet("Empty");

        DimensionRef(wb).Should().Be("A1");
    }

    [Fact]
    public void Dom_Save_Emits_Single_Cell_Ref_For_A_Single_Cell()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["C3"].SetNumber(1);

        DimensionRef(wb).Should().Be("C3");
    }

    [Fact]
    public void Dom_Resave_Corrects_A_Stale_Dimension()
    {
        // The dimension is refreshed from the live extent at Save — editing
        // an opened file must not persist the producer's now-wrong range.
        var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            wb.AddSheet("S")["A1"].SetString("x");
            wb.Save(ms);
        }
        ms.Position = 0;
        using (var wb = Workbook.Open(ms))
        {
            wb["S"]["E9"].SetNumber(2);
            DimensionRef(wb).Should().Be("A1:E9");
        }
    }

    [Fact]
    public void Dimension_Precedes_SheetData_In_The_Part()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetString("x");

        var root = SavedOoxml.SheetXml(wb).Root!;
        var names = root.Elements().Select(e => e.Name.LocalName).ToList();
        names.IndexOf("dimension").Should().BeGreaterThanOrEqualTo(0);
        names.IndexOf("dimension").Should().BeLessThan(names.IndexOf("sheetData"),
            "CT_Worksheet sequences dimension before sheetData");
    }

    [Fact]
    public void Streaming_Save_Emits_Dimension_Covering_The_Used_Range()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.CreateStreaming())
        {
            var s = wb.AddSheet("S");
            s.AppendRow().Set(2, "x");
            s.AppendRow(7).Set(4, 1.0);
            wb.Save(ms, leaveOpen: true);
        }

        var doc = PartFromStream(ms, "xl/worksheets/sheet1.xml");
        var root = doc.Root!;
        root.Element(SavedOoxml.Main + "dimension")?.Attribute("ref")?.Value
            .Should().Be("B1:D7");
        var names = root.Elements().Select(e => e.Name.LocalName).ToList();
        names.IndexOf("dimension").Should().BeLessThan(names.IndexOf("sheetData"));
    }

    [Fact]
    public void Streaming_Save_Emits_A1_For_An_Empty_Sheet()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.CreateStreaming())
        {
            wb.AddSheet("Empty");
            wb.Save(ms, leaveOpen: true);
        }

        PartFromStream(ms, "xl/worksheets/sheet1.xml").Root!
            .Element(SavedOoxml.Main + "dimension")?.Attribute("ref")?.Value
            .Should().Be("A1");
    }

    [Fact]
    public void Streaming_Compressed_TempFiles_Splice_Identically()
    {
        // The splice reads back through the gzip layer — the
        // CompressTempFiles knob must not change the emitted part.
        using var ms = new MemoryStream();
        using (var wb = Workbook.CreateStreaming(new StreamingOptions { CompressTempFiles = true }))
        {
            var s = wb.AddSheet("S");
            s.AppendRow().Set(1, "x");
            wb.Save(ms, leaveOpen: true);
        }

        PartFromStream(ms, "xl/worksheets/sheet1.xml").Root!
            .Element(SavedOoxml.Main + "dimension")?.Attribute("ref")?.Value
            .Should().Be("A1");
    }

    private static System.Xml.Linq.XDocument PartFromStream(MemoryStream ms, string partPath)
    {
        using var zip = new System.IO.Compression.ZipArchive(
            new MemoryStream(ms.ToArray()), System.IO.Compression.ZipArchiveMode.Read);
        var entry = zip.GetEntry(partPath);
        entry.Should().NotBeNull($"part {partPath} must exist");
        using var stream = entry!.Open();
        return System.Xml.Linq.XDocument.Load(stream);
    }
}
