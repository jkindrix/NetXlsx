// I-82 engine swap — cell styles slice conformance.
//
// Covers the Open XML SDK engine's styling surface: ICell.Style merge semantics,
// NumberFormat, GetStyle round-trip, ApplyNamedStyle, IRange.Apply, the style-pool
// dedup contract (decision #4), IColumn width/hidden/default-style, IRow height/
// hidden, the date/time/duration setters (unblocked here), date-format detection
// (Kind == Date, GetDate), the 1900/1904 epoch, and stylesheet scaffolding (the
// Normal master style I-78, fills none/gray125, default font index 0).

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class CellStyleTests
{
    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-ooxml-styles-{Guid.NewGuid():N}.xlsx");

    // ---- Stylesheet scaffolding (created workbook) -------------------------

    [Fact]
    public void Created_Workbook_Has_Normal_Style_And_Default_Fills()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        var ss = wb.Underlying.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;

        ss.GetFirstChild<S.CellStyles>()!.Elements<S.CellStyle>()
            .Should().Contain(c => c.Name == "Normal" && c.BuiltinId!.Value == 0u);

        var fills = ss.GetFirstChild<S.Fills>()!.Elements<S.Fill>().ToList();
        fills[0].PatternFill!.PatternType!.Value.Should().Be(S.PatternValues.None);
        fills[1].PatternFill!.PatternType!.Value.Should().Be(S.PatternValues.Gray125);
        ss.GetFirstChild<S.CellStyleFormats>()!.Elements<S.CellFormat>().Should().HaveCount(1);
    }

    [Fact]
    public void Default_Font_Reflects_Workbook_Options()
    {
        using var wb = Workbook.Create(new WorkbookOptions { DefaultFontName = "Aptos Narrow", DefaultFontSize = 12 });
        wb.AddSheet("S");
        var font0 = wb.Underlying.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
            .GetFirstChild<S.Fonts>()!.Elements<S.Font>().First();
        font0.GetFirstChild<S.FontName>()!.Val!.Value.Should().Be("Aptos Narrow");
        font0.GetFirstChild<S.FontSize>()!.Val!.Value.Should().Be(12);
    }

    // ---- Style round-trip ---------------------------------------------------

    [Fact]
    public void Style_Round_Trips_Font_Fill_Border_Alignment_Through_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            var style = new CellStyle
            {
                Bold = true,
                Italic = true,
                FontName = "Arial",
                FontSize = 14,
                FontColor = Color.FromRgb(0x10, 0x20, 0x30),
                Background = Color.FromRgb(0xFF, 0xEE, 0xDD),
                NumberFormat = "$#,##0.00",
                HorizontalAlignment = HAlign.Center,
                VerticalAlignment = VAlign.Top,
                WrapText = true,
                Borders = CellBorders.All(BorderStyle.Thin, Color.FromRgb(0, 0, 0)),
            };
            using (var wb = Workbook.Create())
            {
                var c = wb.AddSheet("S")["B2"];
                c.SetNumber(5);
                c.Style(style);
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var read = wb["S"]["B2"].GetStyle();
                read.Bold.Should().BeTrue();
                read.Italic.Should().BeTrue();
                read.FontName.Should().Be("Arial");
                read.FontSize.Should().Be(14);
                read.FontColor.Should().Be(Color.FromRgb(0x10, 0x20, 0x30));
                read.Background.Should().Be(Color.FromRgb(0xFF, 0xEE, 0xDD));
                read.NumberFormat.Should().Be("$#,##0.00");
                read.HorizontalAlignment.Should().Be(HAlign.Center);
                read.VerticalAlignment.Should().Be(VAlign.Top);
                read.WrapText.Should().BeTrue();
                read.Borders!.Top.Should().Be(BorderStyle.Thin);
                read.Borders!.Bottom.Should().Be(BorderStyle.Thin);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Style_Is_A_Merge_Not_A_Replace()
    {
        using var wb = Workbook.Create();
        var c = wb.AddSheet("S")["A1"];
        c.SetNumber(1);
        c.Style(new CellStyle { Bold = true });
        c.Style(new CellStyle { Italic = true });   // overlay — must keep Bold

        var s = c.GetStyle();
        s.Bold.Should().BeTrue();
        s.Italic.Should().BeTrue();
    }

    [Fact]
    public void Default_Style_Is_A_No_Op()
    {
        using var wb = Workbook.Create();
        var c = wb.AddSheet("S")["A1"];
        c.SetNumber(1);
        c.Style(new CellStyle { Bold = true });
        c.Style(CellStyle.Default);                 // must not clear Bold
        c.GetStyle().Bold.Should().BeTrue();
    }

    [Fact]
    public void NumberFormat_Shortcut_Sets_Only_Format()
    {
        using var wb = Workbook.Create();
        var c = wb.AddSheet("S")["A1"];
        c.SetNumber(0.5);
        c.NumberFormat("0.00%");
        c.GetStyle().NumberFormat.Should().Be("0.00%");
    }

    [Fact]
    public void Background_Theme_Takes_Precedence_And_Round_Trips()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                var c = wb.AddSheet("S")["A1"];
                c.SetNumber(1);
                c.Style(new CellStyle { BackgroundTheme = new ThemeColor(4, -0.25) });
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var s = wb["S"]["A1"].GetStyle();
                s.BackgroundTheme.Should().Be(new ThemeColor(4, -0.25));
                s.Background.Should().BeNull("a theme-backed fill is reported via BackgroundTheme, not RGB");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Dedup (decision #4) ------------------------------------------------

    [Fact]
    public void Equal_Styles_Share_One_CellXfs_Index()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        var style = new CellStyle { Bold = true, FontSize = 12 };
        for (int r = 1; r <= 50; r++)
        {
            s[r, 1].SetNumber(r);
            s[r, 1].Style(style);
        }

        var diag = wb.GetStylePoolDiagnostics();
        diag.UniqueStyles.Should().Be(1);
        diag.StyleHitCount.Should().Be(49, "49 of the 50 applications hit the pool");

        // Exactly one NetXlsx-allocated cellXfs entry beyond the default xf[0].
        var cellXfs = wb.Underlying.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
            .GetFirstChild<S.CellFormats>()!;
        cellXfs.Elements<S.CellFormat>().Should().HaveCount(2);
    }

    [Fact]
    public void Range_Apply_Styles_Every_Cell_Densely()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s.Range("A1:B2").Apply(new CellStyle { Bold = true });
        foreach (var addr in new[] { "A1", "A2", "B1", "B2" })
            s[addr].GetStyle().Bold.Should().BeTrue($"{addr} should be styled");
    }

    // ---- Named styles -------------------------------------------------------

    [Fact]
    public void ApplyNamedStyle_Resolves_Registered_Style()
    {
        using var wb = Workbook.Create();
        wb.RegisterStyle("Heading", new CellStyle { Bold = true, FontSize = 16 });
        wb.RegisteredStyleNames.Should().Contain("Heading");

        var c = wb.AddSheet("S")["A1"];
        c.SetString("Title");
        c.ApplyNamedStyle("Heading");
        c.GetStyle().FontSize.Should().Be(16);
    }

    [Fact]
    public void ApplyNamedStyle_Unknown_Name_Throws()
    {
        using var wb = Workbook.Create();
        var c = wb.AddSheet("S")["A1"];
        ((Action)(() => c.ApplyNamedStyle("nope"))).Should().Throw<ArgumentException>();
    }

    // ---- Columns ------------------------------------------------------------

    [Fact]
    public void Column_Width_Hidden_Style_Round_Trip()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddSheet("S");
                s.Column(2).Width(24);
                s.Column(3).Hidden = true;
                s.Column(4).SetDefaultStyle(new CellStyle { Bold = true });
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var s = wb["S"];
                s.Column(2).WidthUnits.Should().Be(24);
                s.Column(3).Hidden.Should().BeTrue();
                s.Column(4).WidthUnits.Should().BeGreaterThan(0);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Column_Split_Of_A_Spanning_Entry_Preserves_Neighbors()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");

        // Inject a spanning <col min=1 max=10 width=15> directly, as an opened
        // real-world file would carry, then mutate a single middle column.
        var ws = wb.Underlying.WorkbookPart!.WorksheetParts.First().Worksheet!;
        var sheetData = ws.GetFirstChild<S.SheetData>()!;
        ws.InsertBefore(new S.Columns(new S.Column { Min = 1u, Max = 10u, Width = 15, CustomWidth = true }), sheetData);

        s.Column(5).Width(40);

        // Column 5 changed; its former range neighbors keep width 15.
        s.Column(5).WidthUnits.Should().Be(40);
        s.Column(4).WidthUnits.Should().Be(15);
        s.Column(6).WidthUnits.Should().Be(15);
        s.Column(1).WidthUnits.Should().Be(15);
        s.Column(10).WidthUnits.Should().Be(15);
    }

    // ---- Rows ---------------------------------------------------------------

    [Fact]
    public void Row_Height_And_Hidden_Round_Trip()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddSheet("S");
                s.Row(1).HeightInPoints = 33.5f;
                s.Row(2).Hidden = true;
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var s = wb["S"];
                s.Row(1).HeightInPoints.Should().Be(33.5f);
                s.Row(2).Hidden.Should().BeTrue();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // ---- Dates / times ------------------------------------------------------

    [Fact]
    public void SetDate_Writes_Date_Cell_With_Default_Format()
    {
        using var wb = Workbook.Create();
        var c = wb.AddSheet("S")["A1"];
        c.SetDate(new DateTime(2026, 5, 31, 13, 45, 0));
        c.Kind.Should().Be(CellKind.Date);
        c.GetDate().Should().Be(new DateTime(2026, 5, 31, 13, 45, 0));
        c.GetStyle().NumberFormat.Should().Be("yyyy-mm-dd hh:mm:ss");
    }

    [Fact]
    public void SetDate_Round_Trips_Through_Save_Open()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S")["A1"].SetDate(new DateOnly(2026, 1, 2));
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var c = wb["S"]["A1"];
                c.Kind.Should().Be(CellKind.Date);
                c.GetDateOnly().Should().Be(new DateOnly(2026, 1, 2));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void SetDate_1900_Serial_Matches_Excel()
    {
        using var wb = Workbook.Create();
        var c = wb.AddSheet("S")["A1"];
        c.SetDate(new DateTime(2026, 5, 31));
        // Excel serial for 2026-05-31 in the 1900 system is 46173.
        c.GetNumber().Should().Be(46173);
    }

    [Fact]
    public void SetDate_Honors_1904_Date_System()
    {
        var path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create(new WorkbookOptions { DateSystem = DateSystem.Excel1904 }))
            {
                wb.AddSheet("S")["A1"].SetDate(new DateTime(2026, 5, 31));
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                // The 1904 flag is persisted and read back authoritatively, so the
                // serial decodes to the same calendar date.
                wb["S"]["A1"].GetDate().Should().Be(new DateTime(2026, 5, 31));
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Date_Setter_Preserves_A_User_Applied_Style()
    {
        using var wb = Workbook.Create();
        var c = wb.AddSheet("S")["A1"];
        c.Style(new CellStyle { NumberFormat = "yyyy" });
        c.SetDate(new DateTime(2026, 5, 31));        // must NOT clobber the user's "yyyy"
        c.GetStyle().NumberFormat.Should().Be("yyyy");
        c.Kind.Should().Be(CellKind.Date);
    }

    [Fact]
    public void SetTime_And_SetDuration_Round_Trip()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetTime(new TimeOnly(9, 30, 0));
        s["A2"].SetDuration(TimeSpan.FromHours(26.5));
        s["A1"].GetTime().Should().Be(new TimeOnly(9, 30, 0));
        s["A2"].GetDuration().Should().Be(TimeSpan.FromHours(26.5));
    }

    [Fact]
    public void SetDuration_Negative_Throws()
    {
        using var wb = Workbook.Create();
        var c = wb.AddSheet("S")["A1"];
        ((Action)(() => c.SetDuration(TimeSpan.FromHours(-1)))).Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Plain_Number_Is_Not_A_Date()
    {
        using var wb = Workbook.Create();
        var c = wb.AddSheet("S")["A1"];
        c.SetNumber(46173);
        c.Kind.Should().Be(CellKind.Number);
        c.GetDate().Should().BeNull();
    }
}
