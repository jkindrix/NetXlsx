// I-82 engine swap — CROSS-ENGINE DIFFERENTIAL HARNESS (advisor O-13 / D-11).
//
// The single biggest cutover de-risk. Until now NPOI<->SDK behavioral parity was
// maintained by *hand* cross-checking, with no test that runs the SAME scenario
// through BOTH engines and asserts they agree. This harness does exactly that:
// each case builds a workbook via the public API, round-trips it (Save -> Open),
// reads an observable projection through the public API, and asserts the NPOI
// engine (Workbook.Create()/Open()) and the SDK engine (Workbook.CreateOoxml()/
// OpenOoxml()) produce the SAME observation.
//
// Two deliberate principles (per the slice brief):
//   1. Compare SEMANTICALLY, not byte-identical XML. The engines legitimately
//      differ on cosmetics (element naming, editAs defaults, effective-vs-sparse
//      style materialization). We compare what the public API reports, projected
//      to the axes a scenario exercises.
//   2. Calling the NPOI factories from a test is fine — this is a test, not
//      engine code. The parallel-engine rule forbids NPOI in Internal/Ooxml*.cs,
//      not in the conformance suite.
//
// Malformed-input parity (the load-bearing half — it surfaces the silent-default-
// vs-fail-loud divergence the well-formed paths never exercise) lives in
// CrossEngineMalformedInputTests.cs alongside the I-83 fail-loud alignment.

using System;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class CrossEngineDifferentialTests
{
    // ---- the differential harness -----------------------------------------

    // Runs `build` then `read` through both engines (each via a Save/Open round
    // trip) and returns the two observations for comparison.
    private static (T Npoi, T Sdk) Both<T>(
        Action<IWorkbook> build,
        Func<IWorkbook, T> read,
        WorkbookOptions? options = null)
    {
        T Run(Func<WorkbookOptions?, IWorkbook> create, Func<string, WorkbookOptions?, IWorkbook> open)
        {
            var path = Path.Combine(Path.GetTempPath(), $"netxlsx-diff-{Guid.NewGuid():N}.xlsx");
            try
            {
                using (var wb = create(options)) { build(wb); wb.Save(path); }
                using (var wb = open(path, options)) { return read(wb); }
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }

        var npoi = Run(o => Workbook.Create(o), (p, o) => Workbook.Open(p, o));
        var sdk = Run(o => Workbook.CreateOoxml(o), (p, o) => Workbook.OpenOoxml(p, o));
        return (npoi, sdk);
    }

    // Asserts the two engines agree (structural equivalence for records/tuples).
    private static void AssertAgree<T>((T Npoi, T Sdk) r)
        => r.Sdk.Should().BeEquivalentTo(r.Npoi);

    // ---- cell values + kinds ----------------------------------------------

    private sealed record CellObs(string Str, double? Num, bool? Bool, CellKind Kind);

    private static CellObs Read(ICell c) => new(c.GetString(), c.GetNumber(), c.GetBool(), c.Kind);

    [Fact]
    public void String_Cell_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S")["A1"].SetString("hello world"),
            wb => Read(wb["S"]["A1"])));

    [Fact]
    public void Number_Cell_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S")["A1"].SetNumber(42.5),
            wb => Read(wb["S"]["A1"])));

    [Fact]
    public void Negative_Large_Integer_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S")["A1"].SetNumber(-1234567890123L),
            wb => Read(wb["S"]["A1"])));

    [Fact]
    public void Bool_Cell_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S")["A1"].SetBool(true),
            wb => Read(wb["S"]["A1"])));

    [Fact]
    public void Empty_Cell_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S"),
            wb => Read(wb["S"]["Z9"])));

    [Fact]
    public void Whitespace_Preserving_String_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S")["A1"].SetString("  leading/trailing  "),
            wb => Read(wb["S"]["A1"])));

    [Fact]
    public void Multiple_Cells_And_Rows_Agree()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                s["A1"].SetString("a");
                s["B1"].SetNumber(1);
                s["A2"].SetNumber(2.25);
                s["C3"].SetBool(false);
            },
            wb =>
            {
                var s = wb["S"];
                return new[] { Read(s["A1"]), Read(s["B1"]), Read(s["A2"]), Read(s["C3"]), Read(s["Z9"]) };
            }));

    // ---- dates / time (1900 + 1904 epochs) --------------------------------

    private sealed record DateObs(DateTime? Date, CellKind Kind, bool IsDate);

    [Fact]
    public void Date_1900_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S")["A1"].SetDate(new DateTime(2026, 6, 3, 14, 30, 0)),
            wb =>
            {
                var c = wb["S"]["A1"];
                return new DateObs(c.GetDate(), c.Kind, c.Kind == CellKind.Date);
            }));

    [Fact]
    public void Date_1904_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S")["A1"].SetDate(new DateTime(2026, 6, 3, 14, 30, 0)),
            wb =>
            {
                var c = wb["S"]["A1"];
                return new DateObs(c.GetDate(), c.Kind, c.Kind == CellKind.Date);
            },
            options: new WorkbookOptions { DateSystem = DateSystem.Excel1904 }));

    [Fact]
    public void DateOnly_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S")["A1"].SetDate(new DateOnly(2026, 6, 3)),
            wb =>
            {
                var c = wb["S"]["A1"];
                return new DateObs(c.GetDate(), c.Kind, c.Kind == CellKind.Date);
            }));

    // ---- rich text (incl. the inheriting prefix run, lesson #10) ----------

    // Per-run projection EXCLUDES FontSize on purpose: an unset run size is a
    // documented, by-design engine divergence (NPOI resolves it to the workbook
    // default 11; the SDK faithfully preserves the inherit semantic as null — the
    // SDK-quirk #6 / lesson #10 "no <rPr> means inherit" family). The explicitly
    // sized run is checked separately so explicit sizes are still proven to agree.
    private sealed record RunObs(string Text, bool? Bold, bool? Italic);

    private sealed record RichObs(RunObs[] Runs, double? ExplicitlySizedRun);

    private static RichObs ReadRich(ICell c)
    {
        var rt = c.GetRichText();
        if (rt is null) return new RichObs(Array.Empty<RunObs>(), null);
        var runs = rt.Runs.Select(r => new RunObs(r.Text, r.Style.Bold, r.Style.Italic)).ToArray();
        // The third run carried an explicit FontSize=14; both engines must report it.
        var sized = rt.Runs.Count >= 3 ? rt.Runs[2].Style.FontSize : null;
        return new RichObs(runs, sized);
    }

    [Fact]
    public void RichText_Runs_Agree()
        => AssertAgree(Both(
            wb =>
            {
                var rt = new RichText(
                    new RichTextRun("plain "),                                    // inherits cell font (lesson #10)
                    new RichTextRun("bold", new RichTextStyle { Bold = true }),
                    new RichTextRun(" italic", new RichTextStyle { Italic = true, FontSize = 14 }));
                wb.AddSheet("S")["A1"].SetRichText(rt);
            },
            wb => ReadRich(wb["S"]["A1"])));

    [Fact]
    public void RichText_GetString_Concatenation_Agrees()
        => AssertAgree(Both(
            wb =>
            {
                var rt = new RichText(
                    new RichTextRun("foo"),
                    new RichTextRun("bar", new RichTextStyle { Bold = true }));
                wb.AddSheet("S")["A1"].SetRichText(rt);
            },
            wb => wb["S"]["A1"].GetString()));

    // ---- cell styles -------------------------------------------------------

    // GetStyle() projected to the axes this scenario sets. We do NOT compare the
    // whole CellStyle: the engines legitimately differ on UNSET axes (one may
    // materialize the effective default where the other leaves it null — e.g.
    // NPOI reports NumberFormat "General" / FontSize 11 on cells that set neither,
    // the SDK reports null; SDK-quirk #6). Comparing the set axes is the
    // meaningful parity check; explicit number formats are covered separately.
    private sealed record StyleObs(bool? Bold, double? FontSize, string? FontName,
        HAlign? HAlign, VAlign? VAlign, bool? Wrap);

    private static StyleObs ReadStyle(ICell c)
    {
        var st = c.GetStyle();
        return new StyleObs(st.Bold, st.FontSize, st.FontName,
            st.HorizontalAlignment, st.VerticalAlignment, st.WrapText);
    }

    [Fact]
    public void Font_And_Alignment_Style_Agrees()
        => AssertAgree(Both(
            wb =>
            {
                var c = wb.AddSheet("S")["A1"];
                c.SetString("x");
                c.Style(new CellStyle
                {
                    Bold = true,
                    FontSize = 14,
                    FontName = "Arial",
                    HorizontalAlignment = HAlign.Center,
                    VerticalAlignment = VAlign.Top,
                    WrapText = true,
                });
            },
            wb => ReadStyle(wb["S"]["A1"])));

    [Fact]
    public void NumberFormat_Style_Agrees()
        => AssertAgree(Both(
            wb =>
            {
                var c = wb.AddSheet("S")["A1"];
                c.SetNumber(1234.5);
                c.NumberFormat("#,##0.00");
            },
            wb => wb["S"]["A1"].GetStyle().NumberFormat));

    // ---- merges ------------------------------------------------------------

    [Fact]
    public void Merged_Ranges_Agree()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                s.MergeCells("A1:C3");
                s.MergeCells("E5:F5");
            },
            wb => wb["S"].MergedRanges.OrderBy(r => r, StringComparer.Ordinal).ToArray()));

    // ---- named ranges ------------------------------------------------------

    private sealed record NameObs(string Name, string Formula, string? Scope);

    [Fact]
    public void Named_Ranges_Agree()
        => AssertAgree(Both(
            wb =>
            {
                wb.AddSheet("S");
                wb.AddNamedRange("MyName", "S!$A$1:$B$10");
                wb.AddNamedRange("Scoped", "S!$C$1", sheetScope: "S");
            },
            wb => wb.NamedRanges
                .Select(n => new NameObs(n.Name, n.Formula, n.SheetScope))
                .OrderBy(n => n.Name, StringComparer.Ordinal)
                .ToArray()));

    // ---- sheet-level structure --------------------------------------------

    private sealed record SheetObs(bool Hidden, bool ShowGridlines);

    [Fact]
    public void Sheet_Visibility_And_Gridlines_Agree()
        => AssertAgree(Both(
            wb =>
            {
                wb.AddSheet("First");
                var s2 = wb.AddSheet("Second");
                s2.ShowGridlines = false;
                var s3 = wb.AddSheet("Third");
                s3.Hidden = true;
            },
            wb => Enumerable.Range(0, wb.SheetCount)
                .Select(i => wb[i])
                .Select(s => new SheetObs(s.Hidden, s.ShowGridlines))
                .ToArray()));

    [Fact]
    public void Column_Width_And_Hidden_Agree()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                s.Column("B").Width(24.5);
                s.Column("D").Hidden = true;
            },
            wb =>
            {
                var s = wb["S"];
                return new { B = s.Column("B").WidthUnits, DHidden = s.Column("D").Hidden };
            }));

    [Fact]
    public void Row_Height_And_Hidden_Agree()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                s["A1"].SetString("x");
                s.Row(1).HeightInPoints = 30f;   // Row(int) is 1-based on both engines
                s["A2"].SetString("y");
                s.Row(2).Hidden = true;
            },
            wb =>
            {
                var s = wb["S"];
                return new { H = s.Row(1).HeightInPoints, Hidden = s.Row(2).Hidden };
            }));

    // ---- drawing anchors (geometry parity — lesson #11) -------------------

    // 1x1 PNG.
    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

    private sealed record AnchorObs(string FromCell, string ToCell, int Dx1, int Dy1, int Dx2, int Dy2);

    [Fact]
    public void Picture_TwoCell_Anchor_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S").AddPicture("B2", "D5", Png, ImageFormat.Png),
            wb => wb["S"].Pictures
                .Select(p => new AnchorObs(p.FromCell, p.ToCell, p.Dx1, p.Dy1, p.Dx2, p.Dy2))
                .ToArray()));

    // ---- AutoFilter (incl. the hidden _xlnm._FilterDatabase name) ----------

    [Fact]
    public void AutoFilter_State_And_FilterDatabase_Name_Agree()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                s["A1"].SetString("Region"); s["B1"].SetString("Rev");
                s["A2"].SetString("EU"); s["B2"].SetNumber(100);
                s.SetAutoFilter("A1:B2");
                s.SetAutoFilterColumn(0, FilterCriteria.EqualTo("EU"));
            },
            wb => new
            {
                Has = wb["S"].HasAutoFilter,
                Range = wb["S"].AutoFilterRange,
                Names = wb.NamedRanges.Select(n => $"{n.Name}|{n.Formula}|{n.SheetScope}").ToArray(),
            }));

    [Fact]
    public void Cleared_AutoFilter_State_Agrees()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                s["A1"].SetString("h");
                s.SetAutoFilter("A1:B2");
                s.ClearAutoFilter();
            },
            wb => new
            {
                Has = wb["S"].HasAutoFilter,
                Range = wb["S"].AutoFilterRange,
                // Both engines leave the stale _FilterDatabase name in place.
                NameCount = wb.NamedRanges.Count,
            }));

    // ---- Tables (rich public read-back: name/address/columns/style/totals) --

    [Fact]
    public void Table_With_Totals_Agrees()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                s["A1"].SetString("Region"); s["B1"].SetString("Revenue");
                s["A2"].SetString("EU"); s["B2"].SetNumber(100);
                s["A3"].SetString("US"); s["B3"].SetNumber(200);
                var t = s.AddTable("A1:B3", "Sales", TableStyles.Medium2);
                t.AddTotalsRow();
                t.SetColumnTotalLabel("Region", "Total");
                t.SetColumnTotal("Revenue", TotalsRowFunction.Sum);
            },
            wb =>
            {
                var t = wb["S"].Tables.Single();
                return new
                {
                    t.Name,
                    Display = t.DisplayName,
                    t.Address,
                    Columns = t.ColumnNames.ToArray(),
                    t.HasTotalsRow,
                    Style = t.StyleName,
                    Label = wb["S"]["A4"].GetString(),
                    Formula = wb["S"]["B4"].GetFormula(),
                };
            }));

    // ---- Conditional formatting (count is the only public read-back;
    // emission parity lives in ConditionalFormatTests) -----------------------

    [Fact]
    public void Conditional_Formatting_Count_Agrees()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                for (int r = 1; r <= 5; r++) s[r, 1].SetNumber(r * 10);
                s.AddConditionalFormatting("A1:A5",
                    ConditionalFormat.CellValueGreaterThan("30", new CellStyle { Bold = true }));
                s.AddConditionalFormatting("B1:B5",
                    ConditionalFormat.Formula("ISNUMBER(B1)", new CellStyle { Italic = true }));
                s.RemoveConditionalFormatting(0);
            },
            wb => wb["S"].ConditionalFormattingCount));

    // ---- SortRange (lesson #12: stable, blanks last, numbers before strings) --

    [Fact]
    public void Sorted_Range_Agrees()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                // Mixed kinds + a tie (rows 2/3 tie on column A; column B is
                // the secondary key) + a fully blank row 5 inside the range.
                s["A1"].SetString("B"); s["B1"].SetNumber(2);
                s["A2"].SetString("A"); s["B2"].SetNumber(3);
                s["A3"].SetString("A"); s["B3"].SetNumber(1);
                s["A4"].SetNumber(5); s["B4"].SetBool(true);
                s.SortRange("A1:B5", SortKey.Asc(1), SortKey.Asc(2));
            },
            wb =>
            {
                var s = wb["S"];
                return Enumerable.Range(1, 5)
                    .Select(r => $"{s[r, 1].Kind}|{s[r, 1].GetString()}|{s[r, 2].Kind}|{s[r, 2].GetString()}")
                    .ToArray();
            }));

    // ---- Formulas (formulas/comments/hyperlinks slice) ---------------------

    [Fact]
    public void Formula_Cell_Agrees()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                s["A1"].SetNumber(10);
                s["A2"].SetNumber(20);
                s["A3"].SetFormula("=SUM(A1:A2)");
            },
            wb => (wb["S"]["A3"].GetFormula(), wb["S"]["A3"].Kind)));

    [Fact]
    public void Empty_Formula_Body_Rejection_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S"),
            wb =>
            {
                try { wb["S"]["A1"].SetFormula("="); return "no-throw"; }
                catch (FormulaException) { return "FormulaException"; }
            }));

    [Fact]
    public void Structurally_Broken_Formula_Rejection_Agrees()
        => AssertAgree(Both(
            wb => wb.AddSheet("S"),
            wb =>
            {
                // "=SUM(" is the overlap of NPOI's full parse and the SDK
                // engine's structural validation — both throw FormulaException.
                try { wb["S"]["A1"].SetFormula("=SUM("); return "no-throw"; }
                catch (FormulaException) { return "FormulaException"; }
            }));

    // ---- Comments (formulas/comments/hyperlinks slice) ---------------------

    [Fact]
    public void Comment_Text_And_Author_Agree()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                s["B2"].SetString("data");
                s["B2"].Comment("reviewer flagged");          // default author (I11)
                s["C3"].Comment("second note", "alice");
            },
            wb => (wb["S"]["B2"].GetComment(), wb["S"]["B2"].GetCommentAuthor(),
                   wb["S"]["C3"].GetComment(), wb["S"]["C3"].GetCommentAuthor(),
                   wb["S"]["A1"].GetComment())));

    // ---- Hyperlinks (formulas/comments/hyperlinks slice) -------------------

    [Fact]
    public void Hyperlink_Targets_And_Cell_Text_Agree()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                wb.AddSheet("Other");
                s["D1"].Hyperlink("https://example.com", display: "Example");
                s["D2"].Hyperlink("mailto:foo@example.com");
                s["D3"].Hyperlink("#Other!A1");
                s["D4"].SetString("Click me");
                s["D4"].Hyperlink("HTTPS://Example.COM/MixedCase");
            },
            wb =>
            {
                var s = wb["S"];
                return new[]
                {
                    $"{s["D1"].GetHyperlink()}|{s["D1"].GetString()}",
                    $"{s["D2"].GetHyperlink()}|{s["D2"].GetString()}",
                    $"{s["D3"].GetHyperlink()}",
                    // Verbatim target preservation — no URI canonicalization.
                    $"{s["D4"].GetHyperlink()}|{s["D4"].GetString()}",
                    $"{s["E9"].GetHyperlink() ?? "<null>"}",
                };
            }));

    // ---- Workbook protection (formulas/comments/hyperlinks slice rider) ----

    [Fact]
    public void Workbook_Protection_State_Agrees()
        => AssertAgree(Both(
            wb =>
            {
                wb.AddSheet("S");
                wb.ProtectWithPassword("hunter2",
                    new WorkbookProtection { Structure = true, Windows = true });
            },
            wb => wb.IsProtected));

    [Fact]
    public void Unprotected_Workbook_State_Agrees()
        => AssertAgree(Both(
            wb =>
            {
                wb.AddSheet("S");
                wb.Protect();
                wb.Unprotect();
            },
            wb => wb.IsProtected));

    // ---- date-cell GetString rendering (closeout slice) --------------------
    // The agreement matrix for §7.10 date display-formatting: the SDK engine's
    // ExcelDateFormat renderer vs NPOI's DataFormatter. Formats where NPOI
    // demonstrably mangles (quoted literals, lowercase meridiems, A/P,
    // meridiem-before-hour) are EXCLUDED here and pinned Excel-correct
    // SDK-side in DateGetStringTests, with NPOI's output noted inline.

    private static string DateStr(IWorkbook wb) => wb["S"]["A1"].GetString();

    private static Action<IWorkbook> DateCell(string format, DateTime value) => wb =>
    {
        var c = wb.AddSheet("S")["A1"];
        c.SetDate(value);
        c.NumberFormat(format);
    };

    private static Action<IWorkbook> SerialCell(string format, double serial) => wb =>
    {
        var c = wb.AddSheet("S")["A1"];
        c.SetNumber(serial);
        c.NumberFormat(format);
    };

    [Theory]
    [InlineData("yyyy-mm-dd")]
    [InlineData("yyyy-mm-dd hh:mm:ss")]
    [InlineData("h:mm:ss")]
    [InlineData("m/d/yyyy")]
    [InlineData("m/d/yy")]
    [InlineData("d-mmm-yy")]
    [InlineData("d-mmm")]
    [InlineData("mmm-yy")]
    [InlineData("h:mm AM/PM")]
    [InlineData("h:mm:ss AM/PM")]
    [InlineData("h:mm")]
    [InlineData("m/d/yyyy h:mm")]
    [InlineData("mm:ss")]
    [InlineData("mm:ss.0")]
    [InlineData("dddd, mmmm d, yyyy")]
    [InlineData("ddd d mmmmm yy")]
    [InlineData("yy")]
    [InlineData("hh:mm")]
    [InlineData("hhh:mm")]
    [InlineData("m/d/yyyy;@")]
    [InlineData("[$-409]m/d/yyyy")]
    [InlineData("[Red]yyyy-mm-dd")]
    [InlineData("yyyy\\-mm")]
    [InlineData("dd.mm.yyyy")]
    [InlineData("[h]:mm:ss")]
    [InlineData("[hh]:mm")]
    [InlineData("[mm]:ss")]
    [InlineData("[s]")]
    public void Date_GetString_Afternoon_Agrees(string format)
        => AssertAgree(Both(
            DateCell(format, new DateTime(2026, 6, 3, 13, 45, 30, 500)),
            DateStr));

    [Theory]
    [InlineData("h:mm:ss")]
    [InlineData("hh:mm")]
    [InlineData("h:mm AM/PM")]
    [InlineData("mm:ss.0")]
    public void Date_GetString_Morning_Agrees(string format)
        => AssertAgree(Both(
            DateCell(format, new DateTime(2026, 6, 3, 9, 5, 7)),
            DateStr));

    [Theory]
    [InlineData("[h]:mm:ss")]
    [InlineData("[mm]:ss")]
    [InlineData("[s]")]
    [InlineData("yyyy-mm-dd")]
    public void Date_GetString_Duration_Serial_Agrees(string format)
        => AssertAgree(Both(SerialCell(format, 1.0 + 3.5 / 24.0), DateStr));

    [Theory]
    [InlineData(0.0)]      // the epoch day
    [InlineData(-1.25)]    // negative serial — raw-value fallback on both engines
    public void Date_GetString_Edge_Serials_Agree(double serial)
        => AssertAgree(Both(SerialCell("m/d/yyyy", serial), DateStr));

    [Fact]
    public void Date_GetString_Default_Styles_Agree()
        => AssertAgree(Both(
            wb =>
            {
                var s = wb.AddSheet("S");
                s["A1"].SetDate(new DateOnly(2026, 6, 3));
                s["A2"].SetDate(new DateTime(2026, 6, 3, 9, 5, 7));
                s["A3"].SetTime(new TimeOnly(13, 45, 30));
                s["A4"].SetDuration(TimeSpan.FromHours(27.5));
            },
            wb =>
            {
                var s = wb["S"];
                return new[]
                {
                    s["A1"].GetString(), s["A2"].GetString(),
                    s["A3"].GetString(), s["A4"].GetString(),
                };
            }));

    [Fact]
    public void Date_GetString_1904_Agrees()
        => AssertAgree(Both(
            DateCell("yyyy-mm-dd hh:mm:ss", new DateTime(2026, 6, 3, 13, 45, 30)),
            DateStr,
            options: new WorkbookOptions { DateSystem = DateSystem.Excel1904 }));
}
