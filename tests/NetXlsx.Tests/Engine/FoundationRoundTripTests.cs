// I-82 engine swap — foundation slice conformance.
//
// Exercises the Open XML SDK engine's bones: Create / Open / Save / Dispose,
// AddSheet, sheet enumeration + indexers, and engine discrimination via
// IWorkbook.Underlying. Everything beyond the foundation surface throws
// NotImplementedException on this engine until its slice lands, so these tests
// deliberately stay within Create/AddSheet/Save/Open/enumerate.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class FoundationRoundTripTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void Round_Trips_Empty_Workbook_With_One_Sheet_Through_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb.SheetCount.Should().Be(1);
                wb["S"].Name.Should().Be("S");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Adds_Multiple_Sheets_Preserved_In_Order_Through_Round_Trip()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("Alpha");
                wb.AddSheet("Beta");
                wb.AddSheet("Gamma");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb.SheetCount.Should().Be(3);
                wb[0].Name.Should().Be("Alpha");
                wb[1].Name.Should().Be("Beta");
                wb[2].Name.Should().Be("Gamma");
                wb["Gamma"].Name.Should().Be("Gamma");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Underlying_Exposes_The_Live_SpreadsheetDocument()
    {
        // v2.0.0 (I-82): .Underlying IS the SDK escape hatch (the swap-era
        // OpenXmlDocument member was removed before it ever shipped).
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Underlying.Should().NotBeNull();
        wb.Underlying.WorkbookPart.Should().NotBeNull();
    }

    [Fact]
    public void Cell_Underlying_Materializes_The_Node_And_It_Persists()
    {
        // Pinned at the v2.0.0 cutover (advisor-confirmed Q3): reaching for
        // the raw node is a write-like act — a never-written address
        // materializes its <c> element (decision #40 lazy cells stay lazy
        // until the hatch is used), and the node persists through Save.
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-hatch-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddSheet("S");
                s["B2"].Kind.Should().Be(CellKind.Empty);

                var node = s["B2"].Underlying;
                node.CellReference!.Value.Should().Be("B2");
                wb.Save(path);
            }
            using (var reopened = Workbook.Open(path))
            {
                reopened.Underlying.WorkbookPart!.WorksheetParts.Single().Worksheet!
                    .Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
                    .Should().ContainSingle(c => c.CellReference!.Value == "B2",
                        "the hatch-materialized node persists like any written cell");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Duplicate_Sheet_Name_Throws_Case_Insensitive()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        var act = () => wb.AddSheet("DATA");
        act.Should().Throw<SheetNameException>();
    }

    [Fact]
    public void Invalid_Sheet_Name_Throws()
    {
        using var wb = Workbook.Create();
        var act = () => wb.AddSheet("bad/name");
        act.Should().Throw<SheetNameException>();
    }

    [Fact]
    public void Saves_To_Stream_And_Leaves_It_Open_When_Requested()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        using var ms = new MemoryStream();
        wb.Save(ms, leaveOpen: true);
        ms.CanWrite.Should().BeTrue("the stream must remain usable when leaveOpen is true");
        ms.Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void Opens_From_Stream()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            wb.AddSheet("FromStream");
            wb.Save(ms, leaveOpen: true);
        }
        ms.Position = 0;
        using var reopened = Workbook.Open(ms, leaveOpen: true);
        reopened.SheetCount.Should().Be(1);
        reopened["FromStream"].Name.Should().Be("FromStream");
    }

    [Fact]
    public void Save_Is_Repeatable_From_A_Live_Workbook()
    {
        var first = TempXlsxPath();
        var second = TempXlsxPath();
        try
        {
            using var wb = Workbook.Create();
            wb.AddSheet("S");
            wb.Save(first);
            wb.AddSheet("T");
            wb.Save(second);

            using (var a = Workbook.Open(first))
                a.SheetCount.Should().Be(1);
            using (var b = Workbook.Open(second))
                b.SheetCount.Should().Be(2);
        }
        finally
        {
            if (File.Exists(first)) File.Delete(first);
            if (File.Exists(second)) File.Delete(second);
        }
    }

    [Fact]
    public void Open_Missing_File_Throws_FileNotFound()
    {
        var path = TempXlsxPath();
        var act = () => Workbook.Open(path);
        act.Should().Throw<FileNotFoundException>();
    }

    [Fact]
    public void Open_Malformed_File_Throws_MalformedFileException()
    {
        var path = TempXlsxPath();
        File.WriteAllBytes(path, new byte[] { 0x01, 0x02, 0x03, 0x04, 0x05 });
        try
        {
            var act = () => Workbook.Open(path);
            act.Should().Throw<MalformedFileException>();
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Disposed_Workbook_Member_Access_Throws_ObjectDisposed()
    {
        var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Dispose();

        ((Action)(() => _ = wb.SheetCount)).Should().Throw<ObjectDisposedException>();
        ((Action)(() => _ = wb.Underlying)).Should().Throw<ObjectDisposedException>();
        ((Action)(() => wb.AddSheet("X"))).Should().Throw<ObjectDisposedException>();
    }

    [Fact]
    public void Double_Dispose_Is_Safe()
    {
        var wb = Workbook.Create();
        wb.Dispose();
        var act = () => wb.Dispose();
        act.Should().NotThrow();
    }
}
