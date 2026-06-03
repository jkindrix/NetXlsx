// I-82 engine swap — CROSS-ENGINE DIFFERENTIAL HARNESS: malformed-input parity
// (the load-bearing half — advisor O-13 / D-11, decision I-83).
//
// Well-formed round-trips never exercise the engines' MALFORMED-input paths: the
// default branch only fires on corrupt files. An independent source review found
// the SDK engine SILENTLY DEFAULTED where the NPOI engine FAILS LOUD — a corrupt
// shared-string index read back as "", an unparseable numeric <v> read back as 0,
// a non-integer drawing-anchor marker read back as column 0 (mis-placing the
// drawing). That directly contradicts the library's most-praised property: fail
// loud, nothing silently truncates.
//
// This file feeds hand-corrupted .xlsx files through BOTH engines and asserts they
// agree on failing loud. Where the SDK previously diverged it is now aligned to
// throw MalformedFileException (decision I-83, in Internal/Ooxml*.cs — never the
// NPOI Xssf*.cs). The NPOI engine leaks the raw framework exception (a pre-existing
// trait we don't change); the parity we assert is that NEITHER engine silently
// substitutes a default.
//
// Deliberately-lenient paths that were reviewed and left at parity (both engines
// already agree, so no I-83 change) are pinned at the bottom so a future change
// that makes only one engine throw is caught as a regression.

using System;
using System.IO;
using System.IO.Compression;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class CrossEngineMalformedInputTests
{
    private static string TempPath() => Path.Combine(Path.GetTempPath(), $"netxlsx-malformed-{Guid.NewGuid():N}.xlsx");

    // Builds a real workbook with a shared-string cell (A1) and a numeric cell (A2)
    // via the NPOI engine — which writes a shared-string table — then rewrites
    // xl/worksheets/sheet1.xml through `transform` to inject corruption. NPOI emits
    // `<c r="A1" s="0" t="s"><v>0</v></c>` and `<c r="A2" s="0"><v>42</v></c>`.
    private static string CorruptCellsFile(Func<string, string> transform)
    {
        var path = TempPath();
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("S");
            s["A1"].SetString("hello");
            s["A2"].SetNumber(42.0);
            wb.Save(path);
        }
        RewriteEntry(path, "xl/worksheets/sheet1.xml", transform, required: true);
        return path;
    }

    // Builds a workbook with a two-cell-anchored picture, then corrupts the first
    // drawing-anchor marker (xdr:col) to a non-integer.
    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

    private static string CorruptMarkerFile(string newColValue)
    {
        var path = TempPath();
        using (var wb = Workbook.Create())
        {
            wb.AddSheet("S").AddPicture("B2", "D5", Png, ImageFormat.Png);
            wb.Save(path);
        }
        string? drawingEntry = null;
        using (var z = ZipFile.OpenRead(path))
            foreach (var e in z.Entries)
                if (e.FullName.Contains("drawings/drawing"))
                    drawingEntry = e.FullName;
        drawingEntry.Should().NotBeNull("the picture must produce a drawing part");

        RewriteEntry(path, drawingEntry!, xml =>
        {
            int i = xml.IndexOf(":col>", StringComparison.Ordinal);
            int start = i + ":col>".Length;
            int end = xml.IndexOf("</", start, StringComparison.Ordinal);
            return xml[..start] + newColValue + xml[end..];
        }, required: true);
        return path;
    }

    private static void RewriteEntry(string path, string entryName, Func<string, string> transform, bool required)
    {
        using var z = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = z.GetEntry(entryName);
        if (entry is null) { required.Should().BeFalse($"entry {entryName} not found"); return; }
        string xml;
        using (var r = new StreamReader(entry.Open())) xml = r.ReadToEnd();
        var rewritten = transform(xml);
        rewritten.Should().NotBe(xml, "the corruption transform must actually change the XML");
        entry.Delete();
        var ne = z.CreateEntry(entryName);
        using var w = new StreamWriter(ne.Open());
        w.Write(rewritten);
    }

    // Reading A1 (string) through each engine; both must fail loud.
    private static Action ReadStringNpoi(string p) => () => { using var wb = Workbook.Open(p); _ = wb["S"]["A1"].GetString(); };
    private static Action ReadStringSdk(string p) => () => { using var wb = Workbook.OpenOoxml(p); _ = wb["S"]["A1"].GetString(); };
    private static Action ReadNumberNpoi(string p) => () => { using var wb = Workbook.Open(p); _ = wb["S"]["A2"].GetNumber(); };
    private static Action ReadNumberSdk(string p) => () => { using var wb = Workbook.OpenOoxml(p); _ = wb["S"]["A2"].GetNumber(); };

    // ---- shared-string index corruption -----------------------------------

    [Fact]
    public void SharedString_NonInteger_Index_Fails_Loud_On_Both()
    {
        var path = CorruptCellsFile(x => x.Replace("t=\"s\"><v>0</v>", "t=\"s\"><v>notanumber</v>"));
        try
        {
            ReadStringNpoi(path).Should().Throw<Exception>("the NPOI engine fails loud on a corrupt shared-string index");
            ReadStringSdk(path).Should().Throw<MalformedFileException>("the SDK engine is aligned to fail loud (I-83)");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void SharedString_OutOfRange_Index_Fails_Loud_On_Both()
    {
        var path = CorruptCellsFile(x => x.Replace("t=\"s\"><v>0</v>", "t=\"s\"><v>9999</v>"));
        try
        {
            ReadStringNpoi(path).Should().Throw<Exception>();
            ReadStringSdk(path).Should().Throw<MalformedFileException>();
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void SharedString_Negative_Index_Fails_Loud_On_Both()
    {
        var path = CorruptCellsFile(x => x.Replace("t=\"s\"><v>0</v>", "t=\"s\"><v>-1</v>"));
        try
        {
            ReadStringNpoi(path).Should().Throw<Exception>();
            ReadStringSdk(path).Should().Throw<MalformedFileException>();
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void SharedString_Empty_Index_Fails_Loud_On_Both()
    {
        var path = CorruptCellsFile(x => x.Replace("t=\"s\"><v>0</v>", "t=\"s\"><v></v>"));
        try
        {
            ReadStringNpoi(path).Should().Throw<Exception>();
            ReadStringSdk(path).Should().Throw<MalformedFileException>();
        }
        finally { File.Delete(path); }
    }

    // ---- numeric value corruption -----------------------------------------

    [Fact]
    public void Numeric_Garbage_Value_Fails_Loud_On_Both()
    {
        var path = CorruptCellsFile(x => x.Replace("s=\"0\"><v>42</v>", "s=\"0\"><v>NaNNaN</v>"));
        try
        {
            ReadNumberNpoi(path).Should().Throw<Exception>("the NPOI engine fails loud on a non-numeric numeric cell");
            ReadNumberSdk(path).Should().Throw<MalformedFileException>("the SDK engine is aligned to fail loud (I-83)");
        }
        finally { File.Delete(path); }
    }

    // ---- drawing-anchor marker corruption ---------------------------------

    [Fact]
    public void Anchor_Marker_NonInteger_Fails_Loud_On_Both()
    {
        var path = CorruptMarkerFile("notanumber");
        try
        {
            // NPOI rejects the corrupt anchor on open; the SDK now rejects it when the
            // marker is parsed during Pictures read-back (I-83) instead of silently
            // mis-placing the drawing at column 0.
            Action npoi = () => { using var wb = Workbook.Open(path); foreach (var _ in wb["S"].Pictures) { } };
            Action sdk = () => { using var wb = Workbook.OpenOoxml(path); foreach (var _ in wb["S"].Pictures) { } };
            npoi.Should().Throw<Exception>();
            sdk.Should().Throw<MalformedFileException>();
        }
        finally { File.Delete(path); }
    }

    // ---- autoFilter @ref corruption (slice 7) ------------------------------

    [Fact]
    public void Corrupt_AutoFilter_Ref_Fails_Loud_On_Both()
    {
        // SetAutoFilterColumn parses the existing <autoFilter @ref> to bound
        // the column offset; both engines route the corrupt range through
        // CellAddress.ParseRange and fail loud (SDK-quirk #13 diligence: a new
        // opened-file leaf parse must never silently default).
        var path = TempPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddSheet("S");
                s["A1"].SetString("h");
                s.SetAutoFilter("A1:B2");
                wb.Save(path);
            }
            RewriteEntry(path, "xl/worksheets/sheet1.xml",
                x => x.Replace("ref=\"A1:B2\"", "ref=\"NOT_A_RANGE\""), required: true);

            Action npoi = () =>
            {
                using var wb = Workbook.Open(path);
                wb["S"].SetAutoFilterColumn(0, FilterCriteria.EqualTo("x"));
            };
            Action sdk = () =>
            {
                using var wb = Workbook.OpenOoxml(path);
                wb["S"].SetAutoFilterColumn(0, FilterCriteria.EqualTo("x"));
            };
            npoi.Should().Throw<Exception>();
            sdk.Should().Throw<InvalidCellAddressException>();
        }
        finally { File.Delete(path); }
    }

    // ---- deliberately-lenient paths reviewed and LEFT at parity -----------

    // A corrupt boolean <v> is already at parity: BOTH engines read it as false
    // (neither throws). This was reviewed under I-83 and deliberately not changed —
    // pinned here so a future one-sided change is caught as a regression.
    [Fact]
    public void Corrupt_Bool_Value_Is_Lenient_On_Both()
    {
        var path = TempPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S")["A1"].SetBool(true);
                wb.Save(path);
            }
            RewriteEntry(path, "xl/worksheets/sheet1.xml",
                x => x.Replace("t=\"b\"><v>1</v>", "t=\"b\"><v>xyz</v>"), required: true);

            string npoi, sdk;
            using (var wb = Workbook.Open(path)) npoi = wb["S"]["A1"].GetString();
            using (var wb = Workbook.OpenOoxml(path)) sdk = wb["S"]["A1"].GetString();
            sdk.Should().Be(npoi).And.Be("FALSE");
        }
        finally { File.Delete(path); }
    }

    // ---- annotation corruption (formulas/comments/hyperlinks slice) --------

    private static string AnnotatedFile()
    {
        var path = TempPath();
        using (var wb = Workbook.Create())
        {
            var s = wb.AddSheet("S");
            s["A1"].Hyperlink("https://example.com");
            s["B2"].Comment("note", "alice");
            wb.Save(path);
        }
        return path;
    }

    [Fact]
    public void Dangling_Hyperlink_RelId_Fails_Loud_On_Both()
    {
        // A <hyperlink r:id> with no matching relationship is genuine corruption
        // with no OOXML fallback. NPOI rejects the file at OPEN; the SDK engine
        // rejects it when the link is resolved (I-83) — neither silently drops
        // or substitutes the target.
        var path = AnnotatedFile();
        try
        {
            RewriteEntry(path, "xl/worksheets/sheet1.xml",
                x => x.Replace("<hyperlink ref=\"A1\" r:id=\"rId", "<hyperlink ref=\"A1\" r:id=\"rIdZZ"),
                required: true);

            Action npoi = () => { using var wb = Workbook.Open(path); _ = wb["S"]["A1"].GetHyperlink(); };
            Action sdk = () => { using var wb = Workbook.OpenOoxml(path); _ = wb["S"]["A1"].GetHyperlink(); };
            npoi.Should().Throw<Exception>("NPOI rejects the dangling r:id at open");
            sdk.Should().Throw<MalformedFileException>("the SDK engine fails loud at resolution (I-83)");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Hyperlink_With_Neither_RelId_Nor_Location_Is_Lenient_On_Both()
    {
        // CT_Hyperlink makes both @r:id and @location optional, so the
        // degenerate carrying neither is schema-legal and meaningless — both
        // engines report "no hyperlink" (null) rather than throwing. Reviewed
        // under I-83 and deliberately left lenient; pinned as a regression guard.
        var path = AnnotatedFile();
        try
        {
            RewriteEntry(path, "xl/worksheets/sheet1.xml", x =>
            {
                int i = x.IndexOf("<hyperlink ref=\"A1\"", StringComparison.Ordinal);
                int e = x.IndexOf("/>", i, StringComparison.Ordinal);
                return x[..i] + "<hyperlink ref=\"A1\"" + x[e..];
            }, required: true);

            string? npoi, sdk;
            using (var wb = Workbook.Open(path)) npoi = wb["S"]["A1"].GetHyperlink();
            using (var wb = Workbook.OpenOoxml(path)) sdk = wb["S"]["A1"].GetHyperlink();
            npoi.Should().BeNull();
            sdk.Should().BeNull();
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Comment_AuthorId_OutOfRange_Fails_Loud_On_Both()
    {
        // GetCommentAuthor on an authorId past the <authors> list: NPOI throws
        // ArgumentOutOfRangeException; the SDK engine throws
        // MalformedFileException (I-83). GetComment (the text) stays readable
        // on both — the corruption is confined to the author resolution.
        var path = AnnotatedFile();
        try
        {
            RewriteEntry(path, "xl/comments1.xml",
                x => x.Replace("authorId=\"1\"", "authorId=\"99\""), required: true);

            Action npoi = () => { using var wb = Workbook.Open(path); _ = wb["S"]["B2"].GetCommentAuthor(); };
            Action sdk = () => { using var wb = Workbook.OpenOoxml(path); _ = wb["S"]["B2"].GetCommentAuthor(); };
            npoi.Should().Throw<Exception>();
            sdk.Should().Throw<MalformedFileException>();

            string? npoiText, sdkText;
            using (var wb = Workbook.Open(path)) npoiText = wb["S"]["B2"].GetComment();
            using (var wb = Workbook.OpenOoxml(path)) sdkText = wb["S"]["B2"].GetComment();
            sdkText.Should().Be(npoiText).And.Be("note");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Comment_AuthorId_NonInteger_Fails_Loud_On_Sdk()
    {
        // The one deliberate STRICTNESS divergence in this slice: NPOI's XML
        // deserializer silently coerces a non-integer @authorId to 0 and reports
        // author "" — exactly the silent-default substitution I-83 forbids. The
        // SDK engine fails loud instead. Pinned one-sided (no NPOI assertion on
        // the author value: it is a silent default, not a contract).
        var path = AnnotatedFile();
        try
        {
            RewriteEntry(path, "xl/comments1.xml",
                x => x.Replace("authorId=\"1\"", "authorId=\"xyz\""), required: true);

            Action sdk = () => { using var wb = Workbook.OpenOoxml(path); _ = wb["S"]["B2"].GetCommentAuthor(); };
            sdk.Should().Throw<MalformedFileException>();
        }
        finally { File.Delete(path); }
    }
}
