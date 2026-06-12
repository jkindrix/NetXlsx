// LO interop probe harness (ledger R-25; evidence base for docs/interop.md).
//
//   gen <dir>          - write kitchen.xlsx + streaming.xlsx into <dir>
//   verify <file>      - open <file>, print PROBE lines (exploratory)
//   assert-self <file> - hard-assert our own kitchen output (full contract)
//   assert-lo <file>   - hard-assert a LibreOffice RESAVE of kitchen.xlsx:
//                        the survivor set per docs/interop.md; known LO
//                        drops (autofilter model, workbook structure lock,
//                        theme substitution) are soft-reported, not failed
//   assert-streaming <file>    - hard-assert our own streaming output
//   assert-streaming-lo <file> - same file after an LO resave
//
// Exit code: number of failed assertions (0 = green).

using System;
using System.IO;
using System.Linq;
using NetXlsx;
using ReproTypes;

var mode = args.Length > 0 ? args[0] : "gen";
var target = args.Length > 1 ? args[1] : "/tmp/netxlsx-lo";

switch (mode)
{
    case "gen": Gen(target); break;
    case "verify": Verify(target); break;
    case "assert-self": AssertKitchen(target, loResave: false); break;
    case "assert-lo": AssertKitchen(target, loResave: true); break;
    case "assert-streaming": AssertStreaming(target, loResave: false); break;
    case "assert-streaming-lo": AssertStreaming(target, loResave: true); break;
    default:
        Console.WriteLine("usage: gen <dir> | verify <file> | assert-self <file> | assert-lo <file> | assert-streaming <file> | assert-streaming-lo <file>");
        Environment.Exit(2);
        break;
}
Environment.Exit(Check.Failures);

static byte[] TinyPng() => Convert.FromBase64String(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

static void Gen(string dir)
{
    Directory.CreateDirectory(dir);
    BuildKitchen(Path.Combine(dir, "kitchen.xlsx"));
    BuildStreaming(Path.Combine(dir, "streaming.xlsx"));
    // (The 2026-06-10 review harness also wrote an apostrophe-edge sheet
    // name here to probe LO's silent rename; R-9 made that name invalid
    // at AddSheet, so the probe is retired — the rule is unit-pinned.)
    Console.WriteLine($"GEN done -> {dir}");
}

static void BuildKitchen(string path)
{
    using var wb = Workbook.Create();

    // ---------- Scalars ----------
    var sc = wb.AddSheet("Scalars");
    sc["A1"].SetString("Hello NetXlsx");
    sc["B1"].SetString("héllo 🚀 ünïcode");
    sc["A2"].SetNumber(1234.5678);
    sc["B2"].SetNumber(99.95m); sc["B2"].NumberFormat(NumberFormats.Currency);
    sc["C2"].SetNumber(42);
    sc["D2"].SetNumber(9_876_543_210L);
    sc["A3"].SetDate(new DateTime(2026, 6, 10, 14, 30, 0)); sc["A3"].NumberFormat(NumberFormats.DateTime);
    sc["B3"].SetDate(new DateOnly(2026, 6, 10));
    sc["C3"].SetTime(new TimeOnly(9, 30, 15));
    sc["D3"].SetDuration(new TimeSpan(2, 5, 30));
    sc["A4"].SetBool(true); sc["B4"].SetBool(false);
    sc["A5"].SetFormula("=SUM(A2:A2)*2");
    sc["B5"].SetFormula("=SUM(TwoCells)");
    sc["A6"].SetString("wrapped line one\nline two");
    sc["A6"].Style(new CellStyle { WrapText = true });
    sc.Row(7).HeightInPoints = 30f; sc.Row(7).Set(1, "tall row");
    sc.Row(8).Set(1, "hidden row"); sc.Row(8).Hidden = true;
    sc.Column("E").Width(22);
    sc["F1"].SetString("hidden col"); sc.Column("F").Hidden = true;
    sc.Column("A").AutoSize();
    sc.DefaultColumnWidth = 11.5;
    sc.FreezePane(1, 1);
    wb.AddNamedRange("TwoCells", "Scalars!$A$2:$A$2");
    wb.AddNamedRange("LocalName", "Scalars!$B$2", sheetScope: "Scalars");

    // ---------- Styles ----------
    var st = wb.AddSheet("Styles");
    var brand = new CellStyle
    {
        Bold = true, Italic = true, Underline = UnderlineStyle.Single,
        FontName = "Arial", FontSize = 14,
        FontColor = Color.White, Background = Color.FromHex("#003366"),
        HorizontalAlignment = HAlign.Center, VerticalAlignment = VAlign.Center,
        Borders = CellBorders.All(BorderStyle.Thin, Color.Black),
    };
    st["B2"].SetString("branded"); st["B2"].Style(brand);
    st["C3"].SetString("theme bg"); st["C3"].Style(new CellStyle { BackgroundTheme = new ThemeColor(4) });
    st["D4"].SetString("mixed borders"); st["D4"].Style(new CellStyle
    {
        Borders = new CellBorders(
            Top: BorderStyle.Thick, TopColor: Color.Red,
            Bottom: BorderStyle.Double, BottomColor: Color.Blue),
    });
    wb.RegisterStyle("Brand", brand);
    st["E5"].SetString("named style"); st["E5"].ApplyNamedStyle("Brand");
    st.Range("E7:F8").Value("ns").ApplyNamedStyle("Brand");
    st["B6"].SetString("merged anchor"); st.MergeCells("B6:D6");
    st.MergeCellsStyled("B8:D9", new CellStyle { Background = Color.Yellow, HorizontalAlignment = HAlign.Center });
    st["A10"].SetRichText(new RichText(
        new RichTextRun("Hot ") { Style = new RichTextStyle { Bold = true, Color = Color.Red } },
        new RichTextRun("and") { Style = new RichTextStyle { Italic = true } },
        new RichTextRun(" cold") { Style = new RichTextStyle { Underline = UnderlineStyle.Single, Color = Color.Blue } }));
    st["A12"].SetString("u-single"); st["A12"].Style(new CellStyle { Underline = UnderlineStyle.Single });
    st["A13"].SetString("u-double"); st["A13"].Style(new CellStyle { Underline = UnderlineStyle.Double });
    st["A14"].SetString("u-acct"); st["A14"].Style(new CellStyle { Underline = UnderlineStyle.SingleAccounting });
    st.Range("G1:H3").Value(7).Apply(new CellStyle { Background = Color.LightGray, NumberFormat = NumberFormats.NumberTwo });

    // ---------- Data: table + autofilter + sort ----------
    var da = wb.AddSheet("Data");
    da["A1"].SetString("SKU"); da["B1"].SetString("Price"); da["C1"].SetString("Qty");
    da.Row(2).Set(1, "apple").Set(2, 1.25m).Set(3, 10);
    da.Row(3).Set(1, "banana").Set(2, 0.75m).Set(3, 150);
    da.Row(4).Set(1, "cherry").Set(2, 3.50m).Set(3, 7);
    da.Row(5).Set(1, "date").Set(2, 8.00m).Set(3, 42);
    var tbl = da.AddTable("A1:C5", "Items", TableStyles.Medium2);
    tbl.AddTotalsRow();
    tbl.SetColumnTotal("Price", TotalsRowFunction.Sum);
    tbl.SetColumnTotal("Qty", TotalsRowFunction.Average);
    tbl.SetColumnTotalLabel("SKU", "Total");
    da["E1"].SetString("Name"); da["F1"].SetString("Cat"); da["G1"].SetString("Score");
    da.Row(2).Set(5, "alpha").Set(6, "x").Set(7, 5.0);
    da.Row(3).Set(5, "beta").Set(6, "y").Set(7, 50.0);
    da.Row(4).Set(5, "gamma").Set(6, "x").Set(7, 95.0);
    da.Row(5).Set(5, "delta").Set(6, "z").Set(7, 12.0);
    da.SetAutoFilter("E1:G5");
    da.SetAutoFilterColumn(0, FilterCriteria.Contains("a"));
    da.SetAutoFilterColumn(1, FilterCriteria.In("x", "y"));
    da.SetAutoFilterColumn(2, FilterCriteria.GreaterThan(10).And(FilterCriteria.LessThanOrEqual(100)));
    da["I1"].SetNumber(3); da["I2"].SetNumber(1); da["I3"].SetNumber(4); da["I4"].SetNumber(2);
    da.SortRange("I1:I4", SortKey.Desc(9));
    da.FreezeRows(1);

    // ---------- Visual: CF + DV + comments + hyperlinks + grouping ----------
    var vi = wb.AddSheet("Visual");
    for (int r = 1; r <= 8; r++) { vi[r, 1].SetNumber(r); vi[r, 2].SetNumber(r * 10); vi[r, 3].SetNumber(100 - r * 9); }
    vi.AddConditionalFormatting("A1:A8",
        ConditionalFormat.CellValueGreaterThan("5", new CellStyle { Background = Color.Red, FontColor = Color.White }),
        ConditionalFormat.CellValueBetween("2", "4", new CellStyle { Background = Color.Yellow }),
        ConditionalFormat.Formula("=$A1=3", new CellStyle { Bold = true }));
    vi.AddConditionalFormatting("B1:B8", ConditionalFormat.ColorScale(Color.White, Color.Green));
    vi.AddConditionalFormatting("C1:C8", ConditionalFormat.ColorScale(Color.Red, Color.Yellow, Color.Green));
    vi.AddValidation("E1:E5", DataValidation.List("Red", "Green", "Blue"));
    vi.AddValidation("F1:F5", DataValidation.IntegerBetween(1, 10));
    vi.AddValidation("G1:G3", DataValidation.DateBetween(new DateOnly(2026, 1, 1), new DateOnly(2026, 12, 31)));
    vi.AddValidation("H1:H3", DataValidation.Custom("=H1>0"));
    vi["J1"].SetString("one"); vi["J2"].SetString("two"); vi["J3"].SetString("three");
    vi.AddValidation("K1:K3", DataValidation.ListFromRange("Visual!$J$1:$J$3"));
    vi.AddValidation("I1:I3", DataValidation.TextLengthAtMost(5));
    vi["D1"].Comment("Reviewed by QA", author: "QA Bot");
    vi["D2"].Hyperlink("https://example.com/docs", display: "Docs");
    vi["D3"].Hyperlink("mailto:team@example.com", display: "Mail us");
    vi.Row(12).Set(1, "g1"); vi.Row(13).Set(1, "g2"); vi.Row(14).Set(1, "g3"); vi.Row(15).Set(1, "g4");
    vi.GroupRows(12, 15);
    vi.SetRowGroupCollapsed(12, false);
    vi.GroupColumns(11, 13);

    // ---------- Drawing: pictures + borders + shapes + connectors + charts ----------
    var dr = wb.AddSheet("Drawing");
    var png = TinyPng();
    dr.AddPicture("B2", png);
    var p2 = dr.AddPicture("D2", "F6", png, ImageFormat.Png);
    p2.Border = new PictureBorder { Color = Color.FromHex("#FF0000"), WidthPoints = 2.25 };
    var p3 = dr.AddPicture("D8", "F12", png, ImageFormat.Png);
    p3.Border = new PictureBorder { ThemeColor = new ThemeColor(5), WidthPoints = 1.0 };
    dr.AddShape(ShapeType.Rectangle, "H2", "J5", fillColor: Color.LightGray, lineColor: Color.Black);
    dr.AddShape(ShapeType.Ellipse, "H6", "J9", fillColor: Color.Yellow);
    dr.AddConnector(ConnectorType.Straight, "H11", "J13", lineColor: Color.Red,
        headEnd: ConnectorEnd.Arrow, tailEnd: ConnectorEnd.Oval, lineWidthPoints: 1.5);
    dr["A20"].SetString("Q"); dr["B20"].SetString("Rev");
    dr.Row(21).Set(1, "Q1").Set(2, 100.0);
    dr.Row(22).Set(1, "Q2").Set(2, 140.0);
    dr.Row(23).Set(1, "Q3").Set(2, 90.0);
    dr.Row(24).Set(1, "Q4").Set(2, 180.0);
    dr.AddChart(ChartType.Column, "D20", "K34", "A21:A24", "B21:B24", "Revenue by quarter");
    dr.AddChart(ChartType.Pie, "L20", "R34", "A21:A24", "B21:B24", "Share");

    // ---------- Panes ----------
    var pa = wb.AddSheet("Panes");
    pa["A1"].SetString("split panes sheet"); pa["D10"].SetNumber(1);
    pa.CreateSplitPane(2000, 1500);
    pa.ShowGridlines = false;

    // ---------- Protected ----------
    var pr = wb.AddSheet("Protected");
    pr["A1"].SetString("locked sheet");
    pr.Protect("pw123", SheetProtection.Default);

    // ---------- Typed rows via source generator ----------
    var ty = wb.AddSheet("TypedRows");
    ty["A1"].SetString("SKU"); ty["B1"].SetString("Price"); ty["C1"].SetString("Added");
    ty.AddRows(new[]
    {
        new Product { Sku = "apple", Price = 1.25m, Added = new DateOnly(2026, 1, 5) },
        new Product { Sku = "banana", Price = 0.75m, Added = new DateOnly(2026, 2, 6) },
        new Product { Sku = "cherry", Price = 3.50m, Added = new DateOnly(2026, 3, 7) },
    });

    // ---------- Hidden sheet + workbook protection ----------
    var hi = wb.AddSheet("HiddenSheet");
    hi["A1"].SetString("you can't see me");
    hi.Hidden = true;
    wb.Protect(WorkbookProtection.LockStructure);

    wb.Save(path);
    Console.WriteLine($"kitchen: {new FileInfo(path).Length} bytes, styles={wb.GetStylePoolDiagnostics().UniqueStyles}");
}

static void BuildStreaming(string path)
{
    using var swb = Workbook.CreateStreaming(new StreamingOptions { RowAccessWindowSize = 200 });
    var sh = swb.AddSheet("Big");
    var head = sh.AppendRow();
    head.Set(1, "id").Set(2, "name").Set(3, "ratio").Set(4, "price").Set(5, "when").Set(6, "doubled");
    head.Cell(1).Style(new CellStyle { Bold = true, Background = Color.LightGray });
    for (int i = 1; i <= 5000; i++)
    {
        var row = sh.AppendRow();
        row.Set(1, i).Set(2, $"row-{i}").Set(3, i * 1.5).Set(4, (decimal)i / 4).Set(5, new DateTime(2026, 1, 1).AddHours(i));
        row.Cell(5).NumberFormat(NumberFormats.DateTime);
        row.Cell(6).SetFormula($"=A{row.Index}*2");
    }
    swb.Save(path);
    Console.WriteLine($"streaming: {new FileInfo(path).Length} bytes");
}

static void Verify(string path)
{
    Console.WriteLine($"== verify {Path.GetFileName(path)} ==");
    using var wb = Workbook.Open(path);
    P("sheetCount", wb.SheetCount);
    var names = Enumerable.Range(0, wb.SheetCount).Select(i => wb[i]).Select(s => $"{s.Name}{(s.Hidden ? "(H)" : "")}");
    P("sheets", string.Join("|", names));
    P("wbProtected", wb.IsProtected);
    P("namedRanges", string.Join("|", wb.NamedRanges.Select(n => $"{n.Name}={n.Formula}{(n.SheetScope != null ? $"@{n.SheetScope}" : "")}")));
    P("theme4", wb.ResolveThemeColor(4)?.ToHex() ?? "null");

    if (wb.TryGetSheet("Scalars", out var sc))
    {
        P("sc.A1", sc["A1"].GetString());
        P("sc.B1", sc["B1"].GetString());
        P("sc.A2", sc["A2"].GetNumber());
        P("sc.B2fmt", sc["B2"].GetStyle().NumberFormat ?? "null");
        P("sc.A3", sc["A3"].GetDate()?.ToString("yyyy-MM-dd HH:mm") ?? "null");
        P("sc.C3", sc["C3"].GetTime()?.ToString("HH:mm:ss") ?? "null");
        P("sc.D3", sc["D3"].GetDuration()?.ToString() ?? "null");
        P("sc.A4", sc["A4"].GetBool());
        P("sc.A5f", sc["A5"].GetFormula() ?? "null");
        P("sc.B5f", sc["B5"].GetFormula() ?? "null");
        P("sc.lastRow", sc.LastRowNumber);
        P("sc.row8hidden", sc.Row(8).Hidden);
        P("sc.colFhidden", sc.Column("F").Hidden);
        P("sc.colAwidth", Math.Round(sc.Column("A").WidthUnits, 2));
        P("sc.defColW", sc.DefaultColumnWidth?.ToString() ?? "null");
    }
    if (wb.TryGetSheet("Styles", out var st))
    {
        var b2 = st["B2"].GetStyle();
        P("st.B2", $"bold={b2.Bold} bg={b2.Background?.ToHex()} font={b2.FontName}/{b2.FontSize} u={b2.Underline}");
        P("st.C3themeBg", st["C3"].GetStyle().BackgroundTheme?.Index.ToString() ?? "null");
        P("st.merges", string.Join("|", st.MergedRanges));
        var rt = st["A10"].GetRichText();
        P("st.A10runs", rt is null ? "null" : string.Join("+", rt.Runs.Select(r => r.Text)));
        P("st.E5named", st["E5"].GetStyle().Bold);
    }
    if (wb.TryGetSheet("Data", out var da))
    {
        P("da.tables", string.Join("|", da.Tables.Select(t => $"{t.Name}@{t.Address}tot={t.HasTotalsRow}")));
        P("da.autoFilter", da.AutoFilterRange ?? "null");
        P("da.sortedI", $"{da["I1"].GetNumber()},{da["I2"].GetNumber()},{da["I3"].GetNumber()},{da["I4"].GetNumber()}");
    }
    if (wb.TryGetSheet("Visual", out var vi))
    {
        P("vi.cfCount", vi.ConditionalFormattingCount);
        P("vi.comment", $"{vi["D1"].GetComment()}/by:{vi["D1"].GetCommentAuthor()}");
        P("vi.link", vi["D2"].GetHyperlink() ?? "null");
        P("vi.mailto", vi["D3"].GetHyperlink() ?? "null");
    }
    if (wb.TryGetSheet("Drawing", out var dr))
    {
        P("dr.pictures", dr.Pictures.Count);
        P("dr.borders", string.Join("|", dr.Pictures.Select(p =>
            p.Border is null ? "none" : p.Border.ThemeColor is not null ? $"theme{p.Border.ThemeColor.Index}" : p.Border.Color?.ToHex() ?? "?")));
        P("dr.connectors", dr.Connectors.Count);
    }
    if (wb.TryGetSheet("Protected", out var pr)) P("pr.protected", pr.IsProtected);
    if (wb.TryGetSheet("TypedRows", out var ty))
    {
        var rows = ty.ReadRows().ToList();
        P("ty.rows", $"{rows.Count}:{rows.FirstOrDefault()?.Sku}/{rows.FirstOrDefault()?.Price}/{rows.FirstOrDefault()?.Added}");
    }
    if (wb.TryGetSheet("Big", out var big))
    {
        P("big.lastRow", big.LastRowNumber);
        P("big.B5001", big["B5001"].GetString());
        P("big.F10f", big["F10"].GetFormula() ?? "null");
    }
}

static void P(string key, object? value)
{
    try { Console.WriteLine($"PROBE {key}={value}"); }
    catch (Exception ex) { Console.WriteLine($"PROBE {key}=EXC:{ex.GetType().Name}"); }
}


// ---- Hard assertions -------------------------------------------------------

static void AssertKitchen(string path, bool loResave)
{
    Console.WriteLine($"== assert {(loResave ? "LO-resave" : "self")} {Path.GetFileName(path)} ==");
    using var wb = Workbook.Open(path);

    Check.Eq("sheetCount", wb.SheetCount, 9);
    var names = Enumerable.Range(0, wb.SheetCount).Select(i => wb[i].Name).ToList();
    foreach (var n in new[] { "Scalars", "Styles", "Data", "Visual", "Drawing", "Panes", "Protected", "TypedRows", "HiddenSheet" })
        Check.True($"sheet:{n}", names.Contains(n), "missing sheet");
    Check.True("hiddenSheet", wb["HiddenSheet"].Hidden, "HiddenSheet must stay hidden");

    // Named ranges survive both ways.
    var ranges = wb.NamedRanges.Select(n => n.Name).ToList();
    Check.True("name:TwoCells", ranges.Contains("TwoCells"), "workbook-scope name lost");
    Check.True("name:LocalName", ranges.Contains("LocalName"), "sheet-scope name lost");

    // Workbook structure lock: LO drops it on resave (documented in
    // docs/interop.md) — soft-report so a future LO that preserves it is
    // visible without failing the run.
    if (loResave) Check.Soft("wbProtected(LO drops)", wb.IsProtected);
    else Check.True("wbProtected", wb.IsProtected, "structure lock lost on our own round-trip");

    var sc = wb["Scalars"];
    Check.Eq("sc.A1", sc["A1"].GetString(), "Hello NetXlsx");
    Check.Eq("sc.B1", sc["B1"].GetString(), "héllo 🚀 ünïcode");
    Check.Eq("sc.A2", sc["A2"].GetNumber(), 1234.5678);
    Check.Eq("sc.A4", sc["A4"].GetBool(), true);
    Check.Eq("sc.B4", sc["B4"].GetBool(), false);
    // LO normalizes the single-cell range: SUM(A2:A2) resaves as SUM(A2).
    var a5f = sc["A5"].GetFormula() ?? "";
    Check.True("sc.A5f", a5f.Contains("SUM(A2") && a5f.Contains(")*2"), $"formula was '{a5f}'");
    Check.True("sc.B5f", (sc["B5"].GetFormula() ?? "").Contains("TwoCells"), $"named-range formula was '{sc["B5"].GetFormula()}'");
    Check.Eq("sc.C3", sc["C3"].GetTime()?.ToString("HH:mm:ss"), "09:30:15");
    Check.Eq("sc.A3", sc["A3"].GetDate()?.ToString("yyyy-MM-dd HH:mm"), "2026-06-10 14:30");
    Check.Eq("sc.row8hidden", sc.Row(8).Hidden, true);
    Check.Eq("sc.colFhidden", sc.Column("F").Hidden, true);

    var st = wb["Styles"];
    var b2 = st["B2"].GetStyle();
    Check.Eq("st.B2.bold", b2.Bold, true);
    Check.Eq("st.B2.bg", b2.Background?.ToHex(), "#FF003366");
    Check.True("st.merges.B6:D6", st.MergedRanges.Contains("B6:D6"), $"merges: {string.Join("|", st.MergedRanges)}");
    Check.True("st.merges.B8:D9", st.MergedRanges.Contains("B8:D9"), $"merges: {string.Join("|", st.MergedRanges)}");
    Check.Eq("st.A10richPlain", st["A10"].GetString(), "Hot and cold");
    if (!loResave)
    {
        var rt = st["A10"].GetRichText();
        Check.True("st.A10runs", rt is not null && rt.Runs.Count == 3, "rich-text runs lost on our own round-trip");
    }
    else
    {
        // LO preserves the runs (survivor set) but may restructure them
        // (shared-string conversion); presence is the contract.
        var rt = st["A10"].GetRichText();
        Check.True("st.A10runs(LO)", rt is not null && rt.Runs.Count >= 2, "rich-text formatting lost in LO resave");
    }

    var da = wb["Data"];
    Check.True("da.tableItems", da.Tables.Any(t => t.Name == "Items"), $"tables: {string.Join("|", da.Tables.Select(t => t.Name))}");
    Check.True("da.tableTotals", da.Tables.First(t => t.Name == "Items").HasTotalsRow, "totals row lost");
    if (loResave) Check.Soft("da.autoFilter(LO drops the model)", da.AutoFilterRange ?? "null");
    else Check.Eq("da.autoFilter", da.AutoFilterRange, "E1:G5");
    Check.Eq("da.sortedI", $"{da["I1"].GetNumber()},{da["I2"].GetNumber()},{da["I3"].GetNumber()},{da["I4"].GetNumber()}", "4,3,2,1");

    var vi = wb["Visual"];
    Check.True("vi.cf", vi.ConditionalFormattingCount >= 3, $"CF blocks: {vi.ConditionalFormattingCount}");
    Check.Eq("vi.comment", vi["D1"].GetComment(), "Reviewed by QA");
    Check.Eq("vi.link", vi["D2"].GetHyperlink(), "https://example.com/docs");
    Check.Eq("vi.mailto", vi["D3"].GetHyperlink(), "mailto:team@example.com");

    var dr = wb["Drawing"];
    Check.Eq("dr.pictures", dr.Pictures.Count, 3);
    Check.Eq("dr.connectors", dr.Connectors.Count, 1);
    // LO rewrites the theme-indexed picture border to a literal (R-8
    // evidence) and substitutes its own theme — soft-report both.
    if (loResave)
    {
        Check.Soft("dr.borders(LO rewrites theme-indexed)", string.Join("|", dr.Pictures.Select(pic =>
            pic.Border is null ? "none" : pic.Border.ThemeColor is not null ? $"theme{pic.Border.ThemeColor.Index}" : pic.Border.Color?.ToHex() ?? "?")));
        Check.Soft("theme4(LO substitutes its own theme)", wb.ResolveThemeColor(4)?.ToHex() ?? "null");
    }

    Check.True("pr.protected", wb["Protected"].IsProtected, "sheet protection lost");

    var ty = wb["TypedRows"];
    var rows = ty.ReadRows().ToList();
    Check.Eq("ty.count", rows.Count, 3);
    Check.Eq("ty.first", $"{rows[0].Sku}/{rows[0].Price}/{rows[0].Added:yyyy-MM-dd}", "apple/1.25/2026-01-05");

    Console.WriteLine($"== {(Check.Failures == 0 ? "GREEN" : $"{Check.Failures} FAILURE(S)")} ==");
}

static void AssertStreaming(string path, bool loResave)
{
    Console.WriteLine($"== assert streaming {(loResave ? "LO-resave" : "self")} {Path.GetFileName(path)} ==");
    using var wb = Workbook.Open(path);
    var big = wb["Big"];
    Check.Eq("big.lastRow", big.LastRowNumber, 5001);
    Check.Eq("big.A1", big["A1"].GetString(), "id");
    Check.Eq("big.B5001", big["B5001"].GetString(), "row-5000");
    Check.Eq("big.A5001", big["A5001"].GetNumber(), 5000.0);
    Check.True("big.F10f", (big["F10"].GetFormula() ?? "").Contains("A10*2"), $"formula was '{big["F10"].GetFormula()}'");
    if (loResave)
    {
        // LO calculates on load and caches results — the typed getters
        // must read them (R-7): F10 = A10 * 2, and A10 holds 9.
        Check.Eq("big.F10cached(LO)", big["F10"].GetNumber(), 18.0);
    }
    Console.WriteLine($"== {(Check.Failures == 0 ? "GREEN" : $"{Check.Failures} FAILURE(S)")} ==");
}

internal static class Check
{
    public static int Failures;

    public static void Eq(string key, object? actual, object? expected)
    {
        bool ok = Equals(actual?.ToString(), expected?.ToString());
        Report(key, ok, ok ? actual : $"expected '{expected}' got '{actual}'");
    }

    public static void True(string key, bool condition, string detail)
        => Report(key, condition, condition ? "true" : detail);

    public static void Soft(string key, object? value)
        => Console.WriteLine($"SOFT  {key}={value}");

    private static void Report(string key, bool ok, object? detail)
    {
        Console.WriteLine($"{(ok ? "PASS " : "FAIL ")}{key}={detail}");
        if (!ok) Failures++;
    }
}

namespace ReproTypes
{
    [NetXlsx.Worksheet]
    public partial record Product
    {
        [NetXlsx.Column("SKU", Order = 0)] public string Sku { get; init; } = "";
        [NetXlsx.Column("Price", Order = 1, Format = NetXlsx.NumberFormats.Currency)] public decimal Price { get; init; }
        [NetXlsx.Column("Added", Order = 2)] public System.DateOnly Added { get; init; }
    }
}
