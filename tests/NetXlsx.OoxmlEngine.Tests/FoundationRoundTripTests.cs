// I-82 engine swap — foundation slice conformance.
//
// Exercises the Open XML SDK engine's bones: Create / Open / Save / Dispose,
// AddSheet, sheet enumeration + indexers, and engine discrimination via
// IWorkbook.OpenXmlDocument. Everything beyond the foundation surface throws
// NotImplementedException on this engine until its slice lands, so these tests
// deliberately stay within Create/AddSheet/Save/Open/enumerate.

using System;
using System.IO;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

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
            using (var wb = Workbook.CreateOoxml())
            {
                wb.AddSheet("S");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
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
            using (var wb = Workbook.CreateOoxml())
            {
                wb.AddSheet("Alpha");
                wb.AddSheet("Beta");
                wb.AddSheet("Gamma");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
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
    public void Ooxml_Engine_Exposes_OpenXmlDocument_Non_Null()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("S");
        wb.OpenXmlDocument.Should().NotBeNull();
        wb.OpenXmlDocument!.WorkbookPart.Should().NotBeNull();
    }

    [Fact]
    public void Legacy_Npoi_Engine_Reports_Null_OpenXmlDocument()
    {
        using var wb = Workbook.Create();
        wb.OpenXmlDocument.Should().BeNull();
    }

    [Fact]
    public void Ooxml_Engine_Npoi_Underlying_Hatch_Throws_NotSupported()
    {
        using var wb = Workbook.CreateOoxml();
        var act = () => _ = wb.Underlying;
        act.Should().Throw<NotSupportedException>()
            .WithMessage("*Open XML SDK engine*");
    }

    [Fact]
    public void Duplicate_Sheet_Name_Throws_Case_Insensitive()
    {
        using var wb = Workbook.CreateOoxml();
        wb.AddSheet("Data");
        var act = () => wb.AddSheet("DATA");
        act.Should().Throw<SheetNameException>();
    }

    [Fact]
    public void Invalid_Sheet_Name_Throws()
    {
        using var wb = Workbook.CreateOoxml();
        var act = () => wb.AddSheet("bad/name");
        act.Should().Throw<SheetNameException>();
    }

    [Fact]
    public void Saves_To_Stream_And_Leaves_It_Open_When_Requested()
    {
        using var wb = Workbook.CreateOoxml();
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
        using (var wb = Workbook.CreateOoxml())
        {
            wb.AddSheet("FromStream");
            wb.Save(ms, leaveOpen: true);
        }
        ms.Position = 0;
        using var reopened = Workbook.OpenOoxml(ms, leaveOpen: true);
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
            using var wb = Workbook.CreateOoxml();
            wb.AddSheet("S");
            wb.Save(first);
            wb.AddSheet("T");
            wb.Save(second);

            using (var a = Workbook.OpenOoxml(first))
                a.SheetCount.Should().Be(1);
            using (var b = Workbook.OpenOoxml(second))
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
        var act = () => Workbook.OpenOoxml(path);
        act.Should().Throw<FileNotFoundException>();
    }

    [Fact]
    public void Open_Malformed_File_Throws_MalformedFileException()
    {
        var path = TempXlsxPath();
        File.WriteAllBytes(path, new byte[] { 0x01, 0x02, 0x03, 0x04, 0x05 });
        try
        {
            var act = () => Workbook.OpenOoxml(path);
            act.Should().Throw<MalformedFileException>();
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Disposed_Workbook_Member_Access_Throws_ObjectDisposed()
    {
        var wb = Workbook.CreateOoxml();
        wb.AddSheet("S");
        wb.Dispose();

        ((Action)(() => _ = wb.SheetCount)).Should().Throw<ObjectDisposedException>();
        ((Action)(() => _ = wb.OpenXmlDocument)).Should().Throw<ObjectDisposedException>();
        ((Action)(() => wb.AddSheet("X"))).Should().Throw<ObjectDisposedException>();
    }

    [Fact]
    public void Double_Dispose_Is_Safe()
    {
        var wb = Workbook.CreateOoxml();
        wb.Dispose();
        var act = () => wb.Dispose();
        act.Should().NotThrow();
    }
}
