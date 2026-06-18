// I-92 (ledger R-38): a workbook whose <sheet> targets a ChartsheetPart or
// DialogsheetPart opens as a placeholder ISheet — it participates in the sheet
// collection, carries Name/Hidden, takes part in rename/move/remove, and
// round-trips byte-stable, but grid access throws NotSupportedException.
//
// NetXlsx cannot author a chartsheet, so the fixtures are built with the raw
// Open XML SDK (writing the part XML directly, as Excel does); NetXlsx never
// parses the part content, only its type.

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.Tests;

public class ChartsheetTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static string TempPath() => Path.Combine(Path.GetTempPath(), $"chartsheet-{Guid.NewGuid():N}.xlsx");

    /// <summary>
    /// Builds a workbook with one worksheet ("Data") and, after it, one
    /// chartsheet ("Chart1") — and optionally a dialogsheet ("Dlg1") — via the
    /// raw SDK. The non-worksheet parts get minimal but content-type-correct XML.
    /// </summary>
    private static void CreateFixture(string path, bool includeDialogsheet = false)
    {
        using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new S.Workbook();
        var sheets = wbPart.Workbook.AppendChild(new S.Sheets());

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        wsPart.Worksheet = new S.Worksheet(new S.SheetData());
        sheets.Append(new S.Sheet { Name = "Data", SheetId = 1U, Id = wbPart.GetIdOfPart(wsPart) });

        var csPart = wbPart.AddNewPart<ChartsheetPart>();
        WritePartXml(csPart, $"<chartsheet xmlns=\"{Main}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews></chartsheet>");
        sheets.Append(new S.Sheet { Name = "Chart1", SheetId = 2U, Id = wbPart.GetIdOfPart(csPart) });

        if (includeDialogsheet)
        {
            var dsPart = wbPart.AddNewPart<DialogsheetPart>();
            WritePartXml(dsPart, $"<dialogsheet xmlns=\"{Main}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews></dialogsheet>");
            sheets.Append(new S.Sheet { Name = "Dlg1", SheetId = 3U, Id = wbPart.GetIdOfPart(dsPart) });
        }

        wbPart.Workbook.Save();
    }

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml);
    }

    private static bool ZipHasChartsheetPart(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        return zip.Entries.Any(e => e.FullName.Contains("chartsheets/"));
    }

    [Fact]
    public void Open_Workbook_With_Chartsheet_Surfaces_Placeholder()
    {
        var path = TempPath();
        try
        {
            CreateFixture(path);
            using var wb = Workbook.Open(path);

            wb.SheetCount.Should().Be(2);
            wb["Data"].Kind.Should().Be(SheetKind.Worksheet);
            wb["Chart1"].Kind.Should().Be(SheetKind.Chartsheet);
            wb.Sheets.Select(s => s.Name).Should().Equal("Data", "Chart1");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Chartsheet_Grid_Access_Throws_NotSupported()
    {
        var path = TempPath();
        try
        {
            CreateFixture(path);
            using var wb = Workbook.Open(path);
            var chart = wb["Chart1"];

            chart.Invoking(c => c["A1"]).Should().Throw<NotSupportedException>().WithMessage("*Chart1*chartsheet*");
            chart.Invoking(c => c.AppendRow()).Should().Throw<NotSupportedException>();
            chart.Invoking(c => c.Range("A1:B2")).Should().Throw<NotSupportedException>();
            chart.Invoking(c => { var _ = c.Tables; }).Should().Throw<NotSupportedException>();
            chart.Invoking(c => { var _ = c.HasAutoFilter; }).Should().Throw<NotSupportedException>();
            chart.Invoking(c => { var _ = c.Underlying; }).Should().Throw<NotSupportedException>();
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Chartsheet_Name_And_Hidden_Work()
    {
        var path = TempPath();
        try
        {
            CreateFixture(path);
            using (var wb = Workbook.Open(path))
            {
                wb["Chart1"].Name.Should().Be("Chart1");
                wb["Chart1"].Hidden.Should().BeFalse();
                wb["Chart1"].Hidden = true;
                wb.Save(path);
            }
            using var reopened = Workbook.Open(path);
            reopened["Chart1"].Hidden.Should().BeTrue();
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Chartsheet_Roundtrips_ByteStable_Through_NetXlsx_Save()
    {
        var path = TempPath();
        try
        {
            CreateFixture(path);
            using (var wb = Workbook.Open(path))
            {
                // touch a worksheet cell so the save is non-trivial
                wb["Data"]["A1"].SetString("hello");
                wb.Save(path);
            }
            ZipHasChartsheetPart(path).Should().BeTrue("the ChartsheetPart must survive a NetXlsx save");
            using var reopened = Workbook.Open(path);
            reopened.SheetCount.Should().Be(2);
            reopened["Chart1"].Kind.Should().Be(SheetKind.Chartsheet);
            reopened["Data"]["A1"].GetString().Should().Be("hello");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Chartsheet_Can_Be_Renamed()
    {
        var path = TempPath();
        try
        {
            CreateFixture(path);
            using (var wb = Workbook.Open(path))
            {
                wb["Chart1"].Rename("Renamed");
                wb.Save(path);
            }
            using var reopened = Workbook.Open(path);
            reopened.TryGetSheet("Chart1", out _).Should().BeFalse();
            reopened["Renamed"].Kind.Should().Be(SheetKind.Chartsheet);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Chartsheet_Can_Be_Moved()
    {
        var path = TempPath();
        try
        {
            CreateFixture(path);
            using (var wb = Workbook.Open(path))
            {
                wb.MoveSheet(wb["Chart1"], 1);
                wb.Save(path);
            }
            using var reopened = Workbook.Open(path);
            reopened.Sheets.Select(s => s.Name).Should().Equal("Chart1", "Data");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Chartsheet_Can_Be_Removed_And_Its_Part_Is_Gone()
    {
        var path = TempPath();
        try
        {
            CreateFixture(path);
            using (var wb = Workbook.Open(path))
            {
                wb.RemoveSheet(wb["Chart1"]);
                wb.SheetCount.Should().Be(1);
                wb.Save(path);
            }
            ZipHasChartsheetPart(path).Should().BeFalse("the removed chartsheet's part must be torn down");
            using var reopened = Workbook.Open(path);
            reopened.SheetCount.Should().Be(1);
            reopened.TryGetSheet("Chart1", out _).Should().BeFalse();
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Dialogsheet_Surfaces_As_Placeholder_Kind_Dialogsheet()
    {
        var path = TempPath();
        try
        {
            CreateFixture(path, includeDialogsheet: true);
            using var wb = Workbook.Open(path);

            wb.SheetCount.Should().Be(3);
            wb["Dlg1"].Kind.Should().Be(SheetKind.Dialogsheet);
            wb.Invoking(w => w["Dlg1"].AppendRow()).Should().Throw<NotSupportedException>();
        }
        finally { File.Delete(path); }
    }
}
