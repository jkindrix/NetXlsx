// I-82 engine swap — date-cell GetString rendering (closeout slice).
//
// Per design §7.10 a date-formatted numeric cell's GetString renders through
// the cell's number format; plain numbers stay invariant G17. These tests pin
// the SDK engine's ExcelDateFormat renderer over the oracle-dumped NPOI
// DataFormatter matrix (2026-06-03):
//   - the agreement set (both engines render identically) — also asserted
//     cross-engine in CrossEngineDifferentialTests;
//   - the documented divergence corners where the renderer is Excel-correct
//     and NPOI demonstrably mangles (quoted literals, lowercase meridiems,
//     A/P short meridiems) — pinned here SDK-side only, with the NPOI
//     behavior noted inline so drift in either direction is visible.

using System;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class DateGetStringTests
{
    // 2026-06-03 13:45:30.500 — afternoon, with a half second to lock the
    // truncation-vs-rounding semantics.
    private static readonly DateTime Afternoon = new(2026, 6, 3, 13, 45, 30, 500);
    // 2026-06-03 09:05:07 — single-digit fields exercise pad-vs-no-pad.
    private static readonly DateTime Morning = new(2026, 6, 3, 9, 5, 7);
    // 27.5 elapsed hours — exercises the elapsed [h]/[m]/[s] tokens.
    private const double DurationSerial = 1.0 + 3.5 / 24.0;

    private static string Render(Action<ICell> set, string format)
    {
        using var wb = Workbook.CreateOoxml();
        var c = wb.AddSheet("S")["A1"];
        set(c);
        c.NumberFormat(format);
        return c.GetString();
    }

    private static string RenderDate(DateTime value, string format)
        => Render(c => c.SetDate(value), format);

    private static string RenderSerial(double serial, string format)
        => Render(c => c.SetNumber(serial), format);

    // ---- the agreement set (NPOI DataFormatter renders identically) --------

    [Theory]
    [InlineData("yyyy-mm-dd", "2026-06-03")]
    [InlineData("yyyy-mm-dd hh:mm:ss", "2026-06-03 13:45:30")] // seconds truncate, never round
    [InlineData("h:mm:ss", "13:45:30")]
    [InlineData("m/d/yyyy", "6/3/2026")]
    [InlineData("d-mmm-yy", "3-Jun-26")]
    [InlineData("d-mmm", "3-Jun")]
    [InlineData("mmm-yy", "Jun-26")]
    [InlineData("h:mm AM/PM", "1:45 PM")]
    [InlineData("h:mm:ss AM/PM", "1:45:30 PM")]
    [InlineData("h:mm", "13:45")]
    [InlineData("m/d/yyyy h:mm", "6/3/2026 13:45")]
    [InlineData("mm:ss", "45:30")]            // m before s is minutes
    [InlineData("mm:ss.0", "45:30.5")]        // millisecond fraction, truncated
    [InlineData("dddd, mmmm d, yyyy", "Wednesday, June 3, 2026")]
    [InlineData("ddd d mmmmm yy", "Wed 3 J 26")] // mmmmm = first letter
    [InlineData("yy", "26")]
    [InlineData("y", "26")]                   // 1-2 letters = 2-digit year
    [InlineData("yyy", "2026")]               // 3+ letters = 4-digit year
    [InlineData("hh:mm", "13:45")]
    [InlineData("hhh:mm", "13:45")]           // hour runs clamp at 2
    [InlineData("m/d/yyyy;@", "6/3/2026")]    // dates render the first section
    [InlineData("[$-409]m/d/yyyy", "6/3/2026")] // locale prefix stripped
    [InlineData("[Red]yyyy-mm-dd", "2026-06-03")] // color prefix stripped
    [InlineData("yyyy\\-mm", "2026-06")]      // backslash-escaped literal
    [InlineData("dd.mm.yyyy", "03.06.2026")]  // '.' before a field is a literal
    [InlineData("[h]:mm:ss", "1108237:45:30")] // elapsed hours floor; ss truncates
    [InlineData("[mm]:ss", "66494265:30")]    // elapsed minutes floor
    [InlineData("[s]", "3989655931")]         // a lone elapsed-seconds total ROUNDS (…930.5)
    public void Afternoon_Formats_Render(string format, string expected)
        => RenderDate(Afternoon, format).Should().Be(expected);

    [Theory]
    [InlineData("h:mm:ss", "9:05:07")]        // h unpadded, mm/ss padded
    [InlineData("hh:mm", "09:05")]
    [InlineData("h:mm AM/PM", "9:05 AM")]
    [InlineData("mm:ss.0", "05:07.0")]
    public void Morning_Formats_Render(string format, string expected)
        => RenderDate(Morning, format).Should().Be(expected);

    [Theory]
    [InlineData("[h]:mm:ss", "27:30:00")]
    [InlineData("[mm]:ss", "1650:00")]
    [InlineData("[s]", "99000")]
    [InlineData("yyyy-mm-dd", "1900-01-01")]  // serial 1.1458… resolves into the epoch
    public void Duration_Serial_Formats_Render(string format, string expected)
        => RenderSerial(DurationSerial, format).Should().Be(expected);

    [Fact]
    public void Elapsed_Token_Pads_To_Its_Letter_Count()
    {
        // 0.042361… = 1:01 — [hh] pads, [h] does not.
        RenderSerial(61.0 / 1440.0, "[hh]:mm").Should().Be("01:01");
        RenderSerial(61.0 / 1440.0, "[h]:mm").Should().Be("1:01");
    }

    [Fact]
    public void Serial_Zero_Renders_The_Epoch_Day()
        => RenderSerial(0.0, "m/d/yyyy").Should().Be("12/31/1899");

    [Fact]
    public void Negative_Serial_Falls_Back_To_Raw_G17()
        // There is no date to show for a negative serial (Excel renders
        // ######); both engines return the raw value.
        => RenderSerial(-1.25, "yyyy-mm-dd hh:mm:ss").Should().Be("-1.25");

    [Fact]
    public void Plain_Number_Format_Stays_G17()
    {
        using var wb = Workbook.CreateOoxml();
        var c = wb.AddSheet("S")["A1"];
        c.SetNumber(1234.5);
        c.NumberFormat("#,##0.00");
        c.GetString().Should().Be("1234.5", "non-date number formats are not rendered (§7.10)");
    }

    [Fact]
    public void Default_Date_Styles_Render()
    {
        using var wb = Workbook.CreateOoxml();
        var s = wb.AddSheet("S");
        s["A1"].SetDate(new DateOnly(2026, 6, 3));
        s["A2"].SetDate(Morning);
        s["A3"].SetTime(new TimeOnly(13, 45, 30));
        s["A4"].SetDuration(TimeSpan.FromHours(27.5));
        s["A1"].GetString().Should().Be("2026-06-03");
        s["A2"].GetString().Should().Be("2026-06-03 09:05:07");
        s["A3"].GetString().Should().Be("13:45:30");
        s["A4"].GetString().Should().Be("27:30:00");
    }

    [Fact]
    public void Rendering_Survives_A_Save_Open_Round_Trip()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"date-getstring-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var c = wb.AddSheet("S")["A1"];
                c.SetDate(Afternoon);
                c.NumberFormat("d-mmm-yy");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                wb["S"]["A1"].GetString().Should().Be("3-Jun-26");
            }
        }
        finally { if (System.IO.File.Exists(path)) System.IO.File.Delete(path); }
    }

    [Fact]
    public void Builtin_Date_Format_Id_From_Opened_File_Renders()
    {
        // A cell carrying builtin numFmtId 14 (m/d/yy) — written by this
        // engine via the matching code string — renders through the builtin
        // table after reopen.
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"date-builtin-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var c = wb.AddSheet("S")["A1"];
                c.SetDate(Afternoon);
                c.NumberFormat("m/d/yy");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                wb["S"]["A1"].GetString().Should().Be("6/3/26");
            }
        }
        finally { if (System.IO.File.Exists(path)) System.IO.File.Delete(path); }
    }

    // ---- 1904 epoch ---------------------------------------------------------

    [Fact]
    public void Date_1904_Workbook_Renders_Through_Its_Epoch()
    {
        using var wb = Workbook.CreateOoxml(new WorkbookOptions { DateSystem = DateSystem.Excel1904 });
        var c = wb.AddSheet("S")["A1"];
        c.SetDate(Afternoon);
        c.NumberFormat("yyyy-mm-dd hh:mm:ss");
        c.GetString().Should().Be("2026-06-03 13:45:30");
    }

    // ---- documented divergences (Excel-correct; NPOI mangles) --------------
    // Oracle-dumped 2026-06-03; see the design.md I-82 closeout note. These
    // pin the SDK renderer's Excel-correct output, with NPOI's actual output
    // noted so any future re-alignment is a conscious decision.

    [Fact]
    public void Quoted_Literal_Renders_Verbatim()
        // NPOI renders "2026"26" 6"6"" — it keeps the quote characters and
        // re-interprets the quoted content as format letters.
        => RenderDate(Afternoon, "yyyy\"y\" m\"m\"").Should().Be("2026y 6m");

    [Theory]
    [InlineData("h:mm AM/PM", "1:45 PM")]   // agreement case, for contrast
    [InlineData("h:mm am/pm", "1:45 pm")]   // NPOI: "13:45 a6/p6"
    [InlineData("h:mm A/P", "1:45 P")]      // NPOI: "1:45 PM"
    [InlineData("h:mm a/p", "1:45 p")]      // NPOI: "13:45 a/p"
    [InlineData("AM/PM h:mm", "PM 1:45")]   // NPOI: raw serial fallback
    public void Meridiem_Variants_Render_Excel_Correct(string format, string expected)
        => RenderDate(Afternoon, format).Should().Be(expected);

    [Fact]
    public void Phantom_1900_Leap_Serial_Renders_The_Engines_Epoch_Mapping()
        // Serial 60 is Excel's fictitious 1900-02-29, unrepresentable in
        // DateTime. This engine resolves it to 1900-02-28 (FromSerial's
        // phantom-day subtraction); NPOI rolls it forward to 1900-03-01.
        // Deliberately pinned: both are wrong vs Excel, neither is wronger,
        // and FromSerial's mapping is load-bearing for GetDate.
        => RenderSerial(60.0, "yyyy-mm-dd").Should().Be("1900-02-28");
}
