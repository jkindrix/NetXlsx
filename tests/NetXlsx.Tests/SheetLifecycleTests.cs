// I-90 sheet lifecycle, slice 1 (ledger R-12): ISheet.Rename +
// IWorkbook.MoveSheet per the S2 memo (signed off as amended 2026-06-11).
//
// Rename's document-wide reference rewrite is asserted through the public
// read-back where one exists (GetFormula, NamedRanges, GetHyperlink) and
// through saved-part XML where none does (CF/DV formulas, chart c:f,
// pivot-cache @sheet, sparkline xm:f, table column formulas). Fixtures the
// engine cannot author (pivot caches, sparklines, shared formulas, table
// formulas) are crafted through the public Underlying escape hatch; the
// pivot and sparkline fixtures additionally take the save → reopen path so
// the rewrite is exercised against an OPENED file, per the memo's tests.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xne = DocumentFormat.OpenXml.Office.Excel;

namespace NetXlsx.Tests;

public class SheetLifecycleTests
{
    private static readonly XNamespace Xm = "http://schemas.microsoft.com/office/excel/2006/main";
    private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    /// <summary>
    /// Saves the workbook and returns the XML of the first OPC part whose
    /// path contains <paramref name="pathSubstring"/> — for parts whose
    /// exact name the SDK assigns (charts, tables, pivot caches).
    /// </summary>
    private static XDocument PartContaining(IWorkbook wb, string pathSubstring)
    {
        using var ms = new MemoryStream();
        wb.Save(ms);
        ms.Position = 0;
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var entry = zip.Entries.FirstOrDefault(e =>
                e.FullName.Contains(pathSubstring, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"no saved part matching '{pathSubstring}'");
        using var s = entry.Open();
        return XDocument.Load(s);
    }

    private static IWorkbook SaveAndReopen(IWorkbook wb)
    {
        var ms = new MemoryStream();
        wb.Save(ms);
        wb.Dispose();
        ms.Position = 0;
        return Workbook.Open(ms, leaveOpen: false);
    }

    // ---- Rename: name + lookup coherence --------------------------------

    [Fact]
    public void Rename_Updates_Name_Lookups_And_Saved_Sheet_Entry()
    {
        using var wb = Workbook.Create();
        var alpha = wb.AddSheet("Alpha");
        wb.AddSheet("Beta");

        alpha.Rename("Gamma");

        alpha.Name.Should().Be("Gamma");
        wb["Gamma"].Should().BeSameAs(alpha);
        wb[0].Should().BeSameAs(alpha);
        wb.TryGetSheet("Alpha", out _).Should().BeFalse();
        wb.TryGetSheet("gamma", out var byNewName).Should().BeTrue("sheet lookups are case-insensitive");
        byNewName.Should().BeSameAs(alpha);

        var names = SavedOoxml.WorkbookXml(wb).Root!
            .Element(SavedOoxml.Main + "sheets")!
            .Elements(SavedOoxml.Main + "sheet")
            .Select(e => (string?)e.Attribute("name"));
        names.Should().Equal("Gamma", "Beta");
    }

    [Fact]
    public void Rename_Survives_Save_And_Reopen()
    {
        var wb = Workbook.Create();
        var s = wb.AddSheet("Before");
        s["A1"].SetNumber(7);
        s.Rename("After");

        using var reopened = SaveAndReopen(wb);

        reopened.TryGetSheet("After", out var sheet).Should().BeTrue();
        sheet!["A1"].GetNumber().Should().Be(7);
        reopened.TryGetSheet("Before", out _).Should().BeFalse();
    }

    [Fact]
    public void Rename_To_Exact_Same_Name_Is_NoOp()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("Same");
        s.Rename("Same");
        s.Name.Should().Be("Same");
        wb["Same"].Should().BeSameAs(s);
    }

    [Fact]
    public void Rename_CaseOnly_Change_Is_Allowed_And_Rewrites_References()
    {
        using var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        var calc = wb.AddSheet("Calc");
        calc["A1"].SetFormula("=data!A1+1");

        data.Rename("DATA");

        data.Name.Should().Be("DATA");
        wb["data"].Should().BeSameAs(data, "lookups stay case-insensitive");
        calc["A1"].GetFormula().Should().Be("=DATA!A1+1", "references follow the new casing, as Excel does");
    }

    // ---- Rename: validation ----------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Bad/Name")]
    [InlineData("a:b")]
    [InlineData("History")]
    [InlineData("hIsToRy")]
    [InlineData("'Leading")]
    [InlineData("Trailing'")]
    [InlineData("Ctl\u0001Char")]
    [InlineData("Tab\tChar")]
    [InlineData("ThisNameIsLongerThanThirtyOneChars")]
    public void Rename_Invalid_Name_Throws_And_Leaves_State_Unchanged(string? bad)
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("Keep");

        Action act = () => s.Rename(bad!);

        act.Should().Throw<SheetNameException>();
        s.Name.Should().Be("Keep");
        wb["Keep"].Should().BeSameAs(s);
    }

    [Fact]
    public void Rename_To_Existing_Name_Throws_CaseInsensitive()
    {
        using var wb = Workbook.Create();
        var a = wb.AddSheet("Alpha");
        wb.AddSheet("Beta");
        a["A1"].SetString("x");

        Action act = () => a.Rename("bEtA");

        act.Should().Throw<SheetNameException>();
        a.Name.Should().Be("Alpha");
        wb["Alpha"].Should().BeSameAs(a);
    }

    // ---- Rename: formula rewrite -----------------------------------------

    [Fact]
    public void Rename_Rewrites_Bare_And_Quoted_Formula_References_On_All_Sheets()
    {
        using var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        var calc = wb.AddSheet("Calc");
        data["A1"].SetNumber(21);
        data["B1"].SetFormula("=Data!A1+1");          // self-reference
        calc["A1"].SetFormula("=SUM(Data!A1:A3)*2");   // bare
        calc["A2"].SetFormula("=SUM('Data'!A1)");      // legally over-quoted
        calc["A3"].SetFormula("=Other!A1+Data!B1");    // other sheet names untouched

        data.Rename("Data 2026");

        data["B1"].GetFormula().Should().Be("='Data 2026'!A1+1");
        calc["A1"].GetFormula().Should().Be("=SUM('Data 2026'!A1:A3)*2");
        calc["A2"].GetFormula().Should().Be("=SUM('Data 2026'!A1)");
        calc["A3"].GetFormula().Should().Be("=Other!A1+'Data 2026'!B1");
    }

    [Fact]
    public void Rename_Normalizes_Quoting_To_Bare_When_New_Name_Is_Simple()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("My Data");
        var calc = wb.AddSheet("Calc");
        calc["A1"].SetFormula("='My Data'!A1*3");

        wb["My Data"].Rename("Data2");

        calc["A1"].GetFormula().Should().Be("=Data2!A1*3");
    }

    // The new name is quoted exactly when the bare form would be ambiguous:
    // spaces/punctuation, digit-leading, cell-reference-shaped (A1 and
    // R1C1), boolean literals, embedded apostrophes (doubled).
    [Theory]
    [InlineData("Plain2", "=Plain2!A1")]
    [InlineData("My Sheet", "='My Sheet'!A1")]
    [InlineData("A1", "='A1'!A1")]
    [InlineData("XFD9", "='XFD9'!A1")]
    [InlineData("R1C1", "='R1C1'!A1")]
    [InlineData("RC", "='RC'!A1")]
    [InlineData("TRUE", "='TRUE'!A1")]
    [InlineData("2026", "='2026'!A1")]
    [InlineData("Ver1.2", "='Ver1.2'!A1")]
    [InlineData("O'Brien", "='O''Brien'!A1")]
    public void Rename_Quotes_New_Name_Iff_Needed(string newName, string expected)
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        var calc = wb.AddSheet("Calc");
        calc["A1"].SetFormula("=Data!A1");

        wb["Data"].Rename(newName);

        calc["A1"].GetFormula().Should().Be(expected);
    }

    [Fact]
    public void Rename_Matches_Old_Name_With_Embedded_Apostrophe()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("O'Brien");
        var calc = wb.AddSheet("Calc");
        calc["A1"].SetFormula("=SUM('O''Brien'!A1:A2)");

        wb["O'Brien"].Rename("Ledger");

        calc["A1"].GetFormula().Should().Be("=SUM(Ledger!A1:A2)");
    }

    [Fact]
    public void Rename_Leaves_String_Literals_Untouched_Excel_Parity()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        var calc = wb.AddSheet("Calc");
        calc["A1"].SetFormula("=INDIRECT(\"Data!A1\")");
        calc["A2"].SetFormula("=\"see Data!A1 -> \"&Data!B1");

        wb["Data"].Rename("Moved");

        calc["A1"].GetFormula().Should().Be("=INDIRECT(\"Data!A1\")",
            "Excel does not rewrite string arguments either (documented residual)");
        calc["A2"].GetFormula().Should().Be("=\"see Data!A1 -> \"&Moved!B1",
            "only the reference outside the string literal is rewritten");
    }

    [Fact]
    public void Rename_Of_Sheet_Named_REF_Does_Not_Touch_Error_Literals()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("REF");
        var calc = wb.AddSheet("Calc");
        calc["A1"].SetFormula("=#REF!A1+REF!A1");

        wb["REF"].Rename("Renamed");

        calc["A1"].GetFormula().Should().Be("=#REF!A1+Renamed!A1");
    }

    [Fact]
    public void Rename_Leaves_3D_Span_References_Untouched_Documented_Residual()
    {
        // Excel rewrites 3D endpoints; NetXlsx deliberately does not
        // (documented on ISheet.Rename). This pins the residual so a future
        // change to the lexer is a conscious one.
        using var wb = Workbook.Create();
        wb.AddSheet("Data");
        wb.AddSheet("Zeta");
        var calc = wb.AddSheet("Calc");
        calc["A1"].SetFormula("=SUM(Data:Zeta!A1)");
        calc["A2"].SetFormula("=SUM('Data:Zeta'!A1)");

        wb["Data"].Rename("Q1");

        calc["A1"].GetFormula().Should().Be("=SUM(Data:Zeta!A1)");
        calc["A2"].GetFormula().Should().Be("=SUM('Data:Zeta'!A1)");
    }

    [Fact]
    public void Rename_Rewrites_SharedFormula_Master_Only()
    {
        using var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        var calc = wb.AddSheet("Calc");
        data["A1"].SetNumber(2);

        // The engine writes plain <f>; shared formulas arrive in opened
        // files. Craft the master/follower pair through the escape hatch.
        var master = calc["B1"].Underlying;
        master.CellFormula = new S.CellFormula("Data!A1*2")
        {
            FormulaType = S.CellFormulaValues.Shared,
            SharedIndex = 0U,
            Reference = "B1:B2",
        };
        var follower = calc["B2"].Underlying;
        follower.CellFormula = new S.CellFormula
        {
            FormulaType = S.CellFormulaValues.Shared,
            SharedIndex = 0U,
        };

        wb["Data"].Rename("Data 2026");

        var sheetXml = SavedOoxml.SheetXml(wb, 2);
        var b1 = SavedOoxml.Cell(sheetXml, "B1")!.Element(SavedOoxml.Main + "f")!;
        b1.Value.Should().Be("'Data 2026'!A1*2");
        var b2 = SavedOoxml.Cell(sheetXml, "B2")!.Element(SavedOoxml.Main + "f")!;
        b2.Value.Should().BeEmpty("the follower carries no formula text");
    }

    // ---- Rename: defined names -------------------------------------------

    [Fact]
    public void Rename_Rewrites_DefinedName_Bodies_Including_Xlnm_BuiltIns()
    {
        using var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        wb.AddSheet("Other");
        data["A1"].SetNumber(1);
        wb.AddNamedRange("Total", "Data!$A$1:$A$3");
        wb.AddNamedRange("Local", "'Data'!$B$2", sheetScope: "Data");
        wb.AddNamedRange("_xlnm.Print_Area", "Data!$A$1:$B$5", sheetScope: "Data");
        data.SetAutoFilter("A1:B3"); // writes the hidden _xlnm._FilterDatabase name

        data.Rename("Data 2026");

        var byName = wb.NamedRanges.ToDictionary(n => n.Name, n => n.Formula);
        byName["Total"].Should().Be("'Data 2026'!$A$1:$A$3");
        byName["Local"].Should().Be("'Data 2026'!$B$2");
        byName["_xlnm.Print_Area"].Should().Be("'Data 2026'!$A$1:$B$5",
            "reserved _xlnm.* names must not be filtered out of the rewrite (memo amendment)");
        byName["_xlnm._FilterDatabase"].Should().Be("'Data 2026'!$A$1:$B$3");
    }

    // ---- Rename: hyperlinks ----------------------------------------------

    [Fact]
    public void Rename_Rewrites_Internal_Hyperlink_Locations_And_Leaves_External_Alone()
    {
        using var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        var nav = wb.AddSheet("Nav");
        data["A1"].SetString("target");
        nav["A1"].Hyperlink("#Data!A1", display: "jump");
        nav["A2"].Hyperlink("https://example.com/Data!A1", display: "ext");

        data.Rename("Data 2026");

        nav["A1"].GetHyperlink().Should().Be("'Data 2026'!A1",
            "internal locations follow the rename — this deliberately exceeds Excel");
        nav["A2"].GetHyperlink().Should().Be("https://example.com/Data!A1",
            "external targets are URLs, not sheet references");
    }

    // ---- Rename: CF / DV formulas ----------------------------------------

    [Fact]
    public void Rename_Rewrites_ConditionalFormatting_And_Validation_Formulas()
    {
        using var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        var vis = wb.AddSheet("Visual");
        data["A1"].SetNumber(10);
        vis.AddConditionalFormatting("A1:A8",
            ConditionalFormat.Formula("=Data!$A$1>5", new CellStyle { Bold = true }));
        vis.AddValidation("B1:B3", DataValidation.ListFromRange("Data!$J$1:$J$3"));
        vis.AddValidation("C1:C3", DataValidation.Custom("=C1>Data!$A$1"));

        data.Rename("Data 2026");

        var sheetXml = SavedOoxml.SheetXml(wb, 2).Root!;
        var cfFormulas = sheetXml.Descendants(SavedOoxml.Main + "cfRule")
            .SelectMany(r => r.Elements(SavedOoxml.Main + "formula"))
            .Select(f => f.Value);
        cfFormulas.Should().Contain(f => f.Contains("'Data 2026'!$A$1"));

        var dvFormulas = sheetXml.Descendants(SavedOoxml.Main + "dataValidation")
            .SelectMany(v => v.Elements())
            .Select(f => f.Value)
            .Where(t => t.Length > 0)
            .ToList();
        dvFormulas.Should().Contain("'Data 2026'!$J$1:$J$3");
        dvFormulas.Should().Contain(f => f.Contains("'Data 2026'!$A$1"));
    }

    // ---- Rename: chart series references ----------------------------------

    [Fact]
    public void Rename_Rewrites_Chart_Series_References()
    {
        using var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        for (int r = 1; r <= 4; r++)
        {
            data[r, 1].SetString($"Q{r}");
            data[r, 2].SetNumber(r * 10);
        }
        data.AddChart(ChartType.Column, "D1", "K15", "A1:A4", "B1:B4", "Revenue");

        data.Rename("Data 2026");

        var chartXml = PartContaining(wb, "charts/chart");
        var refs = chartXml.Descendants(ChartNs + "f").Select(f => f.Value).ToList();
        refs.Should().NotBeEmpty();
        refs.Should().OnlyContain(f => f.StartsWith("'Data 2026'!", StringComparison.Ordinal));
    }

    // ---- Rename: opened-file fixtures (pivot cache, sparklines, tables) ---

    [Fact]
    public void Rename_Rewrites_PivotCache_WorksheetSource_On_Opened_File()
    {
        var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        data["A1"].SetString("k"); data["B1"].SetString("v");
        data["A2"].SetString("a"); data["B2"].SetNumber(1);

        // NetXlsx cannot author pivots; craft the cache part through the
        // escape hatch the way an Excel-authored file would carry it.
        var wbPart = wb.Underlying.WorkbookPart!;
        var cachePart = wbPart.AddNewPart<PivotTableCacheDefinitionPart>();
        cachePart.PivotCacheDefinition = new S.PivotCacheDefinition(
            new S.CacheSource(new S.WorksheetSource { Sheet = "Data", Reference = "A1:B2" })
            {
                Type = S.SourceValues.Worksheet,
            },
            new S.CacheFields { Count = 0U });

        using var reopened = SaveAndReopen(wb);
        reopened["Data"].Rename("Data 2026");

        var pivotXml = PartContaining(reopened, "pivotCacheDefinition");
        var source = pivotXml.Descendants(SavedOoxml.Main + "worksheetSource").Single();
        ((string?)source.Attribute("sheet")).Should().Be("Data 2026",
            "a missed pivot-cache rename dangles the cache (memo amendment)");
    }

    [Fact]
    public void Rename_Rewrites_Sparkline_Formulas_On_Opened_File()
    {
        var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        var dash = wb.AddSheet("Dash");
        for (int r = 1; r <= 5; r++) data[r, 1].SetNumber(r);

        // Sparklines live in the worksheet extLst (x14); the engine never
        // authors them — craft what an Excel file would carry.
        dash.Underlying.AppendChild(new S.WorksheetExtensionList(
            new S.WorksheetExtension(
                new X14.SparklineGroups(
                    new X14.SparklineGroup(
                        new X14.Sparklines(
                            new X14.Sparkline(
                                new Xne.Formula("Data!A1:A5"),
                                new Xne.ReferenceSequence("B1"))))))
            {
                Uri = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}",
            }));

        using var reopened = SaveAndReopen(wb);
        reopened["Data"].Rename("Data 2026");

        using var ms = new MemoryStream();
        reopened.Save(ms);
        ms.Position = 0;
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var dashEntry = zip.GetEntry("xl/worksheets/sheet2.xml")!;
        using var s = dashEntry.Open();
        var xmFormulas = XDocument.Load(s).Descendants(Xm + "f").Select(f => f.Value);
        xmFormulas.Should().Contain("'Data 2026'!A1:A5");
    }

    [Fact]
    public void Rename_Rewrites_Table_Column_Formulas()
    {
        using var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        var grid = wb.AddSheet("Grid");
        data["A1"].SetNumber(2);
        grid["A1"].SetString("K"); grid["B1"].SetString("V");
        grid["A2"].SetString("a"); grid["B2"].SetNumber(1);
        grid.AddTable("A1:B2", "Tbl");

        // The engine writes plain table columns; calculated-column and
        // custom totals formulas arrive in opened files — craft them.
        var wsPart = grid.Underlying.WorksheetPart!;
        var columns = wsPart.TableDefinitionParts.Single().Table!
            .GetFirstChild<S.TableColumns>()!.Elements<S.TableColumn>().ToList();
        columns[0].AppendChild(new S.CalculatedColumnFormula("Data!A1*2"));
        columns[1].AppendChild(new S.TotalsRowFormula("SUM(Data!A1:A2)"));

        wb["Data"].Rename("Data 2026");

        var tableXml = PartContaining(wb, "tables/table");
        tableXml.Descendants(SavedOoxml.Main + "calculatedColumnFormula")
            .Single().Value.Should().Be("'Data 2026'!A1*2");
        tableXml.Descendants(SavedOoxml.Main + "totalsRowFormula")
            .Single().Value.Should().Be("SUM('Data 2026'!A1:A2)");
    }

    // ---- MoveSheet: order ---------------------------------------------------

    [Fact]
    public void MoveSheet_To_Front_Middle_And_End_Pins_1Based_Resulting_Position()
    {
        using var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B", "C", "D" }) wb.AddSheet(n);

        wb.MoveSheet(wb["C"], 1);
        Enumerable.Range(0, 4).Select(i => wb[i].Name).Should().Equal("C", "A", "B", "D");

        wb.MoveSheet(wb["C"], 4);
        Enumerable.Range(0, 4).Select(i => wb[i].Name).Should().Equal("A", "B", "D", "C");

        wb.MoveSheet(wb["C"], 2);
        Enumerable.Range(0, 4).Select(i => wb[i].Name).Should().Equal("A", "C", "B", "D");

        var saved = SavedOoxml.WorkbookXml(wb).Root!
            .Element(SavedOoxml.Main + "sheets")!
            .Elements(SavedOoxml.Main + "sheet")
            .Select(e => (string?)e.Attribute("name"));
        saved.Should().Equal("A", "C", "B", "D");
    }

    [Fact]
    public void MoveSheet_To_Current_Position_Is_NoOp()
    {
        using var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B", "C" }) wb.AddSheet(n);

        wb.MoveSheet(wb["B"], 2);

        Enumerable.Range(0, 3).Select(i => wb[i].Name).Should().Equal("A", "B", "C");
    }

    [Fact]
    public void MoveSheet_Order_Survives_Save_And_Reopen()
    {
        var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B", "C" }) wb.AddSheet(n);
        wb["C"]["A1"].SetString("marker");
        wb.MoveSheet(wb["C"], 1);

        using var reopened = SaveAndReopen(wb);

        reopened[0].Name.Should().Be("C");
        reopened[0]["A1"].GetString().Should().Be("marker");
        Enumerable.Range(0, 3).Select(i => reopened[i].Name).Should().Equal("C", "A", "B");
    }

    // ---- MoveSheet: argument contract --------------------------------------

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(4)]
    public void MoveSheet_OutOfRange_Index_Throws(int newIndex)
    {
        using var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B", "C" }) wb.AddSheet(n);

        Action act = () => wb.MoveSheet(wb["B"], newIndex);

        act.Should().Throw<ArgumentOutOfRangeException>();
        Enumerable.Range(0, 3).Select(i => wb[i].Name).Should().Equal("A", "B", "C");
    }

    [Fact]
    public void MoveSheet_Foreign_Sheet_Throws_ArgumentException()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Mine");
        using var other = Workbook.Create();
        var foreign = other.AddSheet("Theirs");

        Action act = () => wb.MoveSheet(foreign, 1);

        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void MoveSheet_Null_Sheet_Throws_ArgumentNullException()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("A");

        Action act = () => wb.MoveSheet(null!, 1);

        act.Should().Throw<ArgumentNullException>();
    }

    // ---- MoveSheet: localSheetId re-index -----------------------------------

    [Fact]
    public void MoveSheet_Reindexes_LocalSheetId_So_Scopes_Track_Their_Sheets()
    {
        using var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B", "C" }) wb.AddSheet(n);
        wb.AddNamedRange("ScopedToA", "A!$A$1", sheetScope: "A");
        wb.AddNamedRange("ScopedToC", "C!$A$1", sheetScope: "C");

        wb.MoveSheet(wb["C"], 1); // order: C, A, B

        // Public read-back resolves localSheetId positionally — the scopes
        // must still report the sheets they were created against.
        var scopes = wb.NamedRanges.ToDictionary(n => n.Name, n => n.SheetScope);
        scopes["ScopedToA"].Should().Be("A");
        scopes["ScopedToC"].Should().Be("C");

        var saved = SavedOoxml.WorkbookXml(wb).Root!
            .Element(SavedOoxml.Main + "definedNames")!
            .Elements(SavedOoxml.Main + "definedName")
            .ToDictionary(e => e.Attribute("name")!.Value, e => (string?)e.Attribute("localSheetId"));
        saved["ScopedToC"].Should().Be("0");
        saved["ScopedToA"].Should().Be("1");
    }

    // ---- MoveSheet: activeTab ------------------------------------------------

    [Fact]
    public void MoveSheet_ActiveTab_Follows_The_Active_Sheet()
    {
        using var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B", "C" }) wb.AddSheet(n);
        // Created workbooks carry no bookViews; craft the view an opened
        // file would have (bookViews precedes <sheets> in CT_Workbook).
        var wbRoot = wb.Underlying.WorkbookPart!.Workbook!;
        wbRoot.InsertBefore(
            new S.BookViews(new S.WorkbookView { ActiveTab = 2U }),
            wbRoot.GetFirstChild<S.Sheets>());

        // The active sheet (C) itself moves to the front.
        wb.MoveSheet(wb["C"], 1);
        ActiveTabOf(wb).Should().Be("0");

        // Another sheet moves past it: C sits at 0; move A (index 1) to the
        // end — C stays at 0, so the attribute is unchanged.
        wb.MoveSheet(wb["A"], 3);
        ActiveTabOf(wb).Should().Be("0");

        // Move B (index 1) to the front: C shifts to 1 and the tab follows.
        wb.MoveSheet(wb["B"], 1);
        ActiveTabOf(wb).Should().Be("1");
    }

    [Fact]
    public void MoveSheet_Clamps_A_Malformed_OutOfRange_ActiveTab()
    {
        using var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B", "C" }) wb.AddSheet(n);
        var wbRoot = wb.Underlying.WorkbookPart!.Workbook!;
        wbRoot.InsertBefore(
            new S.BookViews(new S.WorkbookView { ActiveTab = 9U }),
            wbRoot.GetFirstChild<S.Sheets>());

        wb.MoveSheet(wb["C"], 1); // the clamped active sheet (C, last) moved to front

        ActiveTabOf(wb).Should().Be("0");
    }

    private static string? ActiveTabOf(IWorkbook wb)
        => (string?)SavedOoxml.WorkbookXml(wb).Root!
            .Element(SavedOoxml.Main + "bookViews")!
            .Element(SavedOoxml.Main + "workbookView")!
            .Attribute("activeTab");

    // ---- Combined: rename + move + reopen -----------------------------------

    [Fact]
    public void Rename_Then_Move_Then_Reopen_Keeps_References_And_Order_Coherent()
    {
        var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        var calc = wb.AddSheet("Calc");
        data["A1"].SetNumber(21);
        calc["A1"].SetFormula("=SUM(Data!A1)*2");
        wb.AddNamedRange("T", "Data!$A$1");

        data.Rename("Data 2026");
        wb.MoveSheet(calc, 1);

        using var reopened = SaveAndReopen(wb);

        reopened[0].Name.Should().Be("Calc");
        reopened[1].Name.Should().Be("Data 2026");
        reopened["Calc"]["A1"].GetFormula().Should().Be("=SUM('Data 2026'!A1)*2");
        reopened.NamedRanges.Single(n => n.Name == "T").Formula.Should().Be("'Data 2026'!$A$1");
    }

    // ======================================================================
    // RemoveSheet (I-90 slice 2 / R-12) — delete contract
    // ======================================================================

    private static byte[] SavedBytes(IWorkbook wb)
    {
        using var ms = new MemoryStream();
        wb.Save(ms);
        return ms.ToArray();
    }

    private static List<string> ZipNames(byte[] bytes)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        return zip.Entries.Select(e => e.FullName).ToList();
    }

    /// <summary>
    /// Verifies the memo's zip-level "no orphaned parts or rels" invariant on
    /// a saved package: (a) every internal relationship resolves to an existing
    /// part, and (b) every content part is reachable from the document's
    /// relationship graph (a created-then-edited workbook carries no at-open
    /// orphans, so reachable == all content parts).
    /// </summary>
    private static void AssertNoOrphanPartsOrRels(byte[] bytes)
    {
        using (var ms = new MemoryStream(bytes, writable: false))
        using (var pkg = System.IO.Packaging.Package.Open(ms, FileMode.Open, FileAccess.Read))
        {
            foreach (var part in pkg.GetParts())
            {
                // The .rels parts are infrastructure and cannot themselves
                // carry relationships — skip them.
                if (part.ContentType == "application/vnd.openxmlformats-package.relationships+xml")
                    continue;
                foreach (var rel in part.GetRelationships())
                {
                    if (rel.TargetMode != System.IO.Packaging.TargetMode.Internal) continue;
                    var target = System.IO.Packaging.PackUriHelper.ResolvePartUri(part.Uri, rel.TargetUri);
                    pkg.PartExists(target).Should().BeTrue(
                        $"{part.Uri} relationship {rel.Id} → {rel.TargetUri} must resolve to an existing part");
                }
            }
        }

        using (var ms = new MemoryStream(bytes, writable: false))
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var reachable = new HashSet<string>(
                doc.GetAllParts().Select(p => p.Uri.ToString()), StringComparer.OrdinalIgnoreCase);
            foreach (var name in ZipNames(bytes))
            {
                if (name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)) continue;
                if (name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
                reachable.Should().Contain("/" + name,
                    $"saved part '{name}' must be reachable from the relationship graph (no orphan)");
            }
        }
    }

    // ---- RemoveSheet: lookups / order / guards ---------------------------

    [Fact]
    public void RemoveSheet_Drops_Sheet_From_Lookups_Order_And_Saved_Entry()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("A");
        var b = wb.AddSheet("B");
        wb.AddSheet("C");
        b["A1"].SetString("gone");

        wb.RemoveSheet(b);

        wb.SheetCount.Should().Be(2);
        Enumerable.Range(0, 2).Select(i => wb[i].Name).Should().Equal("A", "C");
        wb.TryGetSheet("B", out _).Should().BeFalse();

        var names = SavedOoxml.WorkbookXml(wb).Root!
            .Element(SavedOoxml.Main + "sheets")!
            .Elements(SavedOoxml.Main + "sheet")
            .Select(e => (string?)e.Attribute("name"));
        names.Should().Equal("A", "C");
    }

    [Fact]
    public void RemoveSheet_Of_The_Only_Sheet_Throws_InvalidOperation()
    {
        using var wb = Workbook.Create();
        var only = wb.AddSheet("Only");

        wb.Invoking(w => w.RemoveSheet(only)).Should().Throw<InvalidOperationException>();
        wb.SheetCount.Should().Be(1, "the rejected removal left the workbook unchanged");
    }

    [Fact]
    public void RemoveSheet_Last_Visible_Throws_But_A_Hidden_Sheet_Is_Removable()
    {
        using var wb = Workbook.Create();
        var visible = wb.AddSheet("Visible");
        var hidden = wb.AddSheet("Hidden");
        hidden.Hidden = true;

        // The last VISIBLE sheet cannot be removed even though a hidden one
        // would remain (a workbook needs >=1 visible sheet).
        wb.Invoking(w => w.RemoveSheet(visible)).Should().Throw<InvalidOperationException>();

        // The hidden sheet can go — a visible one remains.
        wb.RemoveSheet(hidden);
        wb.SheetCount.Should().Be(1);
        wb[0].Should().BeSameAs(visible);
    }

    [Fact]
    public void RemoveSheet_Null_Throws_ArgumentNullException()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("A");
        wb.AddSheet("B");

        wb.Invoking(w => w.RemoveSheet(null!)).Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void RemoveSheet_Foreign_Sheet_Throws_ArgumentException()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Mine");
        wb.AddSheet("Mine2");
        using var other = Workbook.Create();
        var foreign = other.AddSheet("Theirs");

        wb.Invoking(w => w.RemoveSheet(foreign)).Should().Throw<ArgumentException>();
    }

    [Fact]
    public void RemoveSheet_Twice_Throws_ArgumentException_On_The_Stale_Handle()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Keep");
        var doomed = wb.AddSheet("Doomed");

        wb.RemoveSheet(doomed);

        wb.Invoking(w => w.RemoveSheet(doomed))
            .Should().Throw<ArgumentException>("a removed handle is stale, like RemoveTable's precedent");
    }

    // ---- RemoveSheet: removed-wrapper access -----------------------------

    [Fact]
    public void Removed_Sheet_And_Its_Cells_Throw_InvalidOperation_Distinct_From_Disposed()
    {
        using var wb = Workbook.Create();
        var doomed = wb.AddSheet("Doomed");
        wb.AddSheet("Keep");
        var cell = doomed["A1"];
        cell.SetString("x"); // materialize before removal

        wb.RemoveSheet(doomed);

        // The workbook is alive — distinct exception class from a disposed one.
        AssertThrowsRemoved(() => { var _ = doomed.Name; });
        AssertThrowsRemoved(() => { var _ = doomed["B2"]; });
        AssertThrowsRemoved(() => doomed.AppendRow());
        AssertThrowsRemoved(() => doomed.Rename("Whatever"));

        // Cell handles obtained before removal are tombstones too.
        AssertThrowsRemoved(() => { var _ = cell.Address; });
        AssertThrowsRemoved(() => cell.SetString("y"));
        AssertThrowsRemoved(() => { var _ = cell.GetString(); });

        // The surviving sheet is unaffected.
        wb["Keep"]["A1"].SetString("ok");
        wb["Keep"]["A1"].GetString().Should().Be("ok");
    }

    private static void AssertThrowsRemoved(Action act)
        => act.Should().Throw<InvalidOperationException>()
            .Which.Should().NotBeOfType<ObjectDisposedException>(
                "a removed sheet throws InvalidOperationException, not ObjectDisposedException");

    // ---- RemoveSheet: #REF! rewrite --------------------------------------

    [Fact]
    public void RemoveSheet_Rewrites_Bare_And_Quoted_CrossSheet_References_To_RefError()
    {
        using var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        var calc = wb.AddSheet("Calc");
        var other = wb.AddSheet("Other");
        data["A1"].SetNumber(5);
        calc["A1"].SetFormula("=SUM(Data!A1:A3)*2");   // bare
        calc["A2"].SetFormula("=SUM('Data'!A1)");      // quoted
        calc["A3"].SetFormula("=Other!A1+Data!B1");    // Other survives, Data does not
        other["A1"].SetNumber(9);

        wb.RemoveSheet(data);

        calc["A1"].GetFormula().Should().Be("=SUM(#REF!A1:A3)*2");
        calc["A2"].GetFormula().Should().Be("=SUM(#REF!A1)");
        calc["A3"].GetFormula().Should().Be("=Other!A1+#REF!B1");
    }

    [Fact]
    public void RemoveSheet_Rewrites_Internal_Hyperlink_Locations_To_RefError()
    {
        using var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        var nav = wb.AddSheet("Nav");
        data["A1"].SetString("target");
        nav["A1"].Hyperlink("#Data!A1", display: "jump");
        nav["A2"].Hyperlink("https://example.com/Data!A1", display: "ext");

        wb.RemoveSheet(data);

        nav["A1"].GetHyperlink().Should().Be("#REF!A1",
            "internal locations to a removed sheet rewrite to #REF! (consistent with rename touching them)");
        nav["A2"].GetHyperlink().Should().Be("https://example.com/Data!A1",
            "external URLs are not sheet references");
    }

    // ---- RemoveSheet: defined names --------------------------------------

    [Fact]
    public void RemoveSheet_Purges_Scoped_Names_Reindexes_Later_Scopes_And_RefErrors_Bodies()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("A");
        var data = wb.AddSheet("Data");
        wb.AddSheet("C");
        wb.AddNamedRange("Total", "Data!$A$1:$A$3");                 // workbook-scoped body
        wb.AddNamedRange("ScopedToData", "Data!$B$2", sheetScope: "Data");
        wb.AddNamedRange("ScopedToC", "C!$A$1", sheetScope: "C");

        wb.RemoveSheet(data);

        var names = wb.NamedRanges.ToDictionary(n => n.Name, n => n);
        names.Should().NotContainKey("ScopedToData", "names scoped to the removed sheet are deleted");
        names["ScopedToC"].SheetScope.Should().Be("C", "later sheet-scopes re-index positionally");
        names["Total"].Formula.Should().Be("#REF!$A$1:$A$3", "a body referencing the removed sheet is #REF!'d");

        // C was at index 2 (A=0, Data=1, C=2); after the shrink it is index 1.
        var saved = SavedOoxml.WorkbookXml(wb).Root!
            .Element(SavedOoxml.Main + "definedNames")!
            .Elements(SavedOoxml.Main + "definedName")
            .ToDictionary(e => e.Attribute("name")!.Value, e => (string?)e.Attribute("localSheetId"));
        saved["ScopedToC"].Should().Be("1");
    }

    // ---- RemoveSheet: activeTab clamp ------------------------------------

    [Fact]
    public void RemoveSheet_Clamps_ActiveTab_Into_The_Shrunken_Range()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("A");
        wb.AddSheet("B");
        var c = wb.AddSheet("C");
        var wbRoot = wb.Underlying.WorkbookPart!.Workbook!;
        wbRoot.InsertBefore(
            new S.BookViews(new S.WorkbookView { ActiveTab = 2U }),
            wbRoot.GetFirstChild<S.Sheets>());

        wb.RemoveSheet(c); // tab 2 now out of range -> clamped to 1

        ((string?)SavedOoxml.WorkbookXml(wb).Root!
            .Element(SavedOoxml.Main + "bookViews")!
            .Element(SavedOoxml.Main + "workbookView")!
            .Attribute("activeTab")).Should().Be("1");
    }

    // ---- RemoveSheet: part cleanup ---------------------------------------

    [Fact]
    public void RemoveSheet_Cleans_Descendant_Parts_With_No_Orphans_Or_Dangling_Rels()
    {
        using var wb = Workbook.Create();
        var rich = wb.AddSheet("Rich");
        wb.AddSheet("Keep");
        rich["A1"].SetString("Cat"); rich["B1"].SetString("Val"); // string headers (AddTable requires them)
        for (int r = 2; r <= 4; r++) { rich[r, 1].SetString($"Q{r}"); rich[r, 2].SetNumber(r * 10); }
        rich.AddChart(ChartType.Column, "D1", "K15", "A2:A4", "B2:B4", "Rev"); // drawing + chart parts
        rich["A1"].Comment("note");                                            // comments part
        rich.AddTable("A1:B4", "Tbl");                                         // table part

        var before = ZipNames(SavedBytes(wb));
        before.Should().Contain(n => n.Contains("drawings/", StringComparison.OrdinalIgnoreCase));
        before.Should().Contain(n => n.Contains("tables/", StringComparison.OrdinalIgnoreCase));

        wb.RemoveSheet(rich);

        var after = SavedBytes(wb);
        AssertNoOrphanPartsOrRels(after);
        ZipNames(after).Should().NotContain(n => n.Contains("tables/", StringComparison.OrdinalIgnoreCase),
            "the removed sheet's table part is gone");
    }

    [Fact]
    public void RemoveSheet_Drops_CalcChain_Wholesale()
    {
        using var wb = Workbook.Create();
        var a = wb.AddSheet("A");
        var b = wb.AddSheet("B");
        a["A1"].SetNumber(1);

        // The engine never authors calcChain; craft one an opened file carries.
        var wbPart = wb.Underlying.WorkbookPart!;
        var ccPart = wbPart.AddNewPart<CalculationChainPart>();
        ccPart.CalculationChain = new S.CalculationChain(
            new S.CalculationCell { CellReference = "A1", SheetId = 1 });

        ZipNames(SavedBytes(wb)).Should().Contain(n => n.Contains("calcChain", StringComparison.OrdinalIgnoreCase),
            "the crafted calcChain is present before removal");

        wb.RemoveSheet(b);

        ZipNames(SavedBytes(wb)).Should().NotContain(n => n.Contains("calcChain", StringComparison.OrdinalIgnoreCase),
            "calcChain is dropped wholesale on removal (c/@i is a sheetId, not a position)");
    }

    [Fact]
    public void RemoveSheet_Removes_PivotCaches_Sourced_From_It_On_Opened_File()
    {
        var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        wb.AddSheet("Keep");
        data["A1"].SetString("k"); data["B1"].SetString("v");
        data["A2"].SetString("a"); data["B2"].SetNumber(1);

        var wbPart = wb.Underlying.WorkbookPart!;
        var cachePart = wbPart.AddNewPart<PivotTableCacheDefinitionPart>();
        cachePart.PivotCacheDefinition = new S.PivotCacheDefinition(
            new S.CacheSource(new S.WorksheetSource { Sheet = "Data", Reference = "A1:B2" })
            {
                Type = S.SourceValues.Worksheet,
            },
            new S.CacheFields { Count = 0U });

        using var reopened = SaveAndReopen(wb);
        reopened.RemoveSheet(reopened["Data"]);

        var after = SavedBytes(reopened);
        ZipNames(after).Should().NotContain(
            n => n.Contains("pivotCacheDefinition", StringComparison.OrdinalIgnoreCase),
            "a cache sourced from the removed sheet is deleted, not left dangling");
        AssertNoOrphanPartsOrRels(after);
    }

    [Fact]
    public void RemoveSheet_Survives_Save_And_Reopen()
    {
        var wb = Workbook.Create();
        var data = wb.AddSheet("Data");
        var calc = wb.AddSheet("Calc");
        data["A1"].SetNumber(7);
        calc["A1"].SetFormula("=Data!A1+1");

        wb.RemoveSheet(data);

        using var reopened = SaveAndReopen(wb);
        reopened.SheetCount.Should().Be(1);
        reopened.TryGetSheet("Data", out _).Should().BeFalse();
        reopened["Calc"]["A1"].GetFormula().Should().Be("=#REF!A1+1");
    }

    // ---- AddSheet(name, index) + Sheets (S19, §6.2 sketch reconciliation) ----

    [Fact]
    public void AddSheet_At_Index_Inserts_At_OneBased_Position()
    {
        using var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B", "C" }) wb.AddSheet(n);

        var front = wb.AddSheet("Front", 1);
        front.Name.Should().Be("Front");
        wb[0].Should().BeSameAs(front);
        Enumerable.Range(0, 4).Select(i => wb[i].Name).Should().Equal("Front", "A", "B", "C");

        wb.AddSheet("Mid", 3);
        Enumerable.Range(0, 5).Select(i => wb[i].Name).Should().Equal("Front", "A", "Mid", "B", "C");
    }

    [Fact]
    public void AddSheet_At_Index_Append_Position_Equals_Plain_AddSheet()
    {
        using var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B" }) wb.AddSheet(n);

        // SheetCount + 1 is the append position — identical to AddSheet(name).
        var tail = wb.AddSheet("Tail", wb.SheetCount + 1);
        wb[wb.SheetCount - 1].Should().BeSameAs(tail);
        Enumerable.Range(0, 3).Select(i => wb[i].Name).Should().Equal("A", "B", "Tail");
    }

    [Theory]
    [InlineData(0)]   // below the 1-based floor
    [InlineData(5)]   // SheetCount (3) + 2 — past the append position
    public void AddSheet_At_Index_OutOfRange_Throws_And_Adds_Nothing(int index)
    {
        using var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B", "C" }) wb.AddSheet(n);

        Action act = () => wb.AddSheet("X", index);
        act.Should().Throw<ArgumentOutOfRangeException>();
        wb.SheetCount.Should().Be(3, "a rejected index must not add a sheet");
    }

    [Fact]
    public void AddSheet_At_Index_Invalid_Name_Throws_And_Adds_Nothing()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("A");

        Action dup = () => wb.AddSheet("A", 1);
        dup.Should().Throw<SheetNameException>();
        wb.SheetCount.Should().Be(1);
        wb[0].Name.Should().Be("A");
    }

    [Fact]
    public void AddSheet_At_Index_Order_Survives_Save_And_Reopen()
    {
        using var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B", "C" }) wb.AddSheet(n);
        wb.AddSheet("Inserted", 2);

        using var reopened = SaveAndReopen(wb);
        Enumerable.Range(0, 4).Select(i => reopened[i].Name).Should().Equal("A", "Inserted", "B", "C");
    }

    [Fact]
    public void Sheets_Matches_Indexer_And_SheetCount_In_TabOrder()
    {
        using var wb = Workbook.Create();
        foreach (var n in new[] { "A", "B", "C" }) wb.AddSheet(n);

        wb.Sheets.Should().HaveCount(wb.SheetCount);
        wb.Sheets.Select(s => s.Name).Should().Equal("A", "B", "C");
        for (int i = 0; i < wb.SheetCount; i++)
            wb.Sheets[i].Should().BeSameAs(wb[i]);
    }

    [Fact]
    public void Sheets_Includes_Hidden_Sheets()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Visible");
        wb.AddSheet("Hidden").Hidden = true;

        wb.Sheets.Select(s => s.Name).Should().Equal("Visible", "Hidden");
    }

    [Fact]
    public void Sheets_Is_Empty_For_New_Workbook()
    {
        using var wb = Workbook.Create();
        wb.Sheets.Should().BeEmpty();
    }

    [Fact]
    public void Sheets_Is_A_Snapshot_Not_A_Live_View()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("A");

        var snapshot = wb.Sheets;
        snapshot.Should().HaveCount(1);

        wb.AddSheet("B");
        snapshot.Should().HaveCount(1, "an already-returned list is a snapshot");
        wb.Sheets.Should().HaveCount(2, "a fresh call reflects the new sheet");
    }
}
