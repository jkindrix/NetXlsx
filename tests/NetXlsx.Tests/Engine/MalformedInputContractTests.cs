// Malformed-input fail-loud contract (decision I-83).
//
// Well-formed round-trips never exercise the engine's MALFORMED-input paths: the
// default branch only fires on corrupt files. These tests feed hand-corrupted
// .xlsx files through the engine and pin the fail-loud contract: a corrupt value
// read from an opened file throws MalformedFileException — it is NEVER silently
// substituted with a default ("", 0, column 0). Deliberately-lenient paths that
// were reviewed under I-83 and left lenient are pinned at the bottom so a future
// one-sided change is caught as a regression.
//
// Until the v2.0.0 cutover this file was CrossEngineMalformedInputTests, feeding
// the same corrupt files through BOTH engines and asserting same-behavior. The
// cross-engine de-risk mission completed at the flip (A1 disposition (b)): the
// NPOI half is gone, and the SDK-side MalformedFileException pins — which were
// always the load-bearing assertions — remain, now against the default factories.

using System;
using System.IO;
using System.IO.Compression;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class MalformedInputContractTests
{
    private static string TempPath() => Path.Combine(Path.GetTempPath(), $"netxlsx-malformed-{Guid.NewGuid():N}.xlsx");

    // Builds a real workbook with a string cell (A1) and a numeric cell (A2),
    // then rewrites xl/worksheets/sheet1.xml through `transform` to inject
    // corruption. The engine writes A1 as an INLINE string — the shared-string
    // corruption cases swap the whole <c> span for a shared-string cell with a
    // corrupt <v> index (no sharedStrings part is needed: resolution fails loud
    // before/at the table lookup either way).
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

    // Replaces the whole <x:c r="..."> ... </x:c> (or self-closed) span of one
    // cell. The engine serializes worksheet parts with the x: prefix.
    private static string ReplaceCellSpan(string xml, string a1, string replacement)
    {
        int start = xml.IndexOf($"r=\"{a1}\"", StringComparison.Ordinal);
        start.Should().BeGreaterThan(0, $"cell {a1} must exist in the sheet XML");
        start = xml.LastIndexOf('<', start);
        int selfClose = xml.IndexOf("/>", start, StringComparison.Ordinal);
        int close = xml.IndexOf("</x:c>", start, StringComparison.Ordinal);
        int end = close >= 0 && (selfClose < 0 || close < selfClose)
            ? close + "</x:c>".Length
            : selfClose + "/>".Length;
        return xml[..start] + replacement + xml[end..];
    }

    private static string SharedStringCellFile(string corruptIndex) =>
        CorruptCellsFile(x => ReplaceCellSpan(x, "A1", $"<x:c r=\"A1\" t=\"s\"><x:v>{corruptIndex}</x:v></x:c>"));

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

    private static Action ReadString(string p) => () => { using var wb = Workbook.Open(p); _ = wb["S"]["A1"].GetString(); };
    private static Action ReadNumber(string p) => () => { using var wb = Workbook.Open(p); _ = wb["S"]["A2"].GetNumber(); };

    // ---- shared-string index corruption -----------------------------------

    [Fact]
    public void SharedString_NonInteger_Index_Fails_Loud()
    {
        var path = SharedStringCellFile("notanumber");
        try { ReadString(path).Should().Throw<MalformedFileException>("a corrupt shared-string index must never read back as \"\" (I-83)"); }
        finally { File.Delete(path); }
    }

    [Fact]
    public void SharedString_OutOfRange_Index_Fails_Loud()
    {
        var path = SharedStringCellFile("9999");
        try { ReadString(path).Should().Throw<MalformedFileException>(); }
        finally { File.Delete(path); }
    }

    [Fact]
    public void SharedString_Negative_Index_Fails_Loud()
    {
        var path = SharedStringCellFile("-1");
        try { ReadString(path).Should().Throw<MalformedFileException>(); }
        finally { File.Delete(path); }
    }

    [Fact]
    public void SharedString_Empty_Index_Fails_Loud()
    {
        var path = SharedStringCellFile("");
        try { ReadString(path).Should().Throw<MalformedFileException>(); }
        finally { File.Delete(path); }
    }

    // ---- numeric value corruption -----------------------------------------

    [Fact]
    public void Numeric_Garbage_Value_Fails_Loud()
    {
        var path = CorruptCellsFile(x => ReplaceCellSpan(x, "A2", "<x:c r=\"A2\"><x:v>NaNNaN</x:v></x:c>"));
        try { ReadNumber(path).Should().Throw<MalformedFileException>("a non-numeric numeric <v> must never read back as 0 (I-83)"); }
        finally { File.Delete(path); }
    }

    // ---- drawing-anchor marker corruption ---------------------------------

    [Fact]
    public void Anchor_Marker_NonInteger_Fails_Loud()
    {
        // A non-integer marker must reject at Pictures read-back (I-83), not
        // silently mis-place the drawing at column 0.
        var path = CorruptMarkerFile("notanumber");
        try
        {
            Action read = () => { using var wb = Workbook.Open(path); foreach (var _ in wb["S"].Pictures) { } };
            read.Should().Throw<MalformedFileException>();
        }
        finally { File.Delete(path); }
    }

    // ---- autoFilter @ref corruption (slice 7) ------------------------------

    [Fact]
    public void Corrupt_AutoFilter_Ref_Fails_Loud()
    {
        // SetAutoFilterColumn parses the existing <autoFilter @ref> to bound
        // the column offset; the corrupt range routes through
        // CellAddress.ParseRange and fails loud (SDK-quirk #13 diligence: an
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

            Action act = () =>
            {
                using var wb = Workbook.Open(path);
                wb["S"].SetAutoFilterColumn(0, FilterCriteria.EqualTo("x"));
            };
            act.Should().Throw<InvalidCellAddressException>();
        }
        finally { File.Delete(path); }
    }

    // ---- deliberately-lenient paths reviewed and LEFT lenient --------------

    // A corrupt boolean <v> reads back as false (no throw). Reviewed under
    // I-83 and deliberately not changed — pinned here so a future change that
    // makes it throw is a conscious decision, not an accident.
    [Fact]
    public void Corrupt_Bool_Value_Is_Lenient()
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
                x => x.Replace("t=\"b\"><x:v>1</x:v>", "t=\"b\"><x:v>xyz</x:v>"), required: true);

            using var opened = Workbook.Open(path);
            opened["S"]["A1"].GetString().Should().Be("FALSE");
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
    public void Dangling_Hyperlink_RelId_Fails_Loud()
    {
        // A <hyperlink r:id> with no matching relationship is genuine corruption
        // with no OOXML fallback — the engine rejects it at resolution (I-83)
        // rather than silently dropping or substituting the target.
        var path = AnnotatedFile();
        try
        {
            RewriteEntry(path, "xl/worksheets/sheet1.xml",
                x => x.Replace("<x:hyperlink ref=\"A1\" r:id=\"", "<x:hyperlink ref=\"A1\" r:id=\"ZZ"),
                required: true);

            Action act = () => { using var wb = Workbook.Open(path); _ = wb["S"]["A1"].GetHyperlink(); };
            act.Should().Throw<MalformedFileException>();
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Hyperlink_With_Neither_RelId_Nor_Location_Is_Lenient()
    {
        // CT_Hyperlink makes both @r:id and @location optional, so the
        // degenerate carrying neither is schema-legal and meaningless — the
        // engine reports "no hyperlink" (null) rather than throwing. Reviewed
        // under I-83 and deliberately left lenient; pinned as a regression guard.
        var path = AnnotatedFile();
        try
        {
            RewriteEntry(path, "xl/worksheets/sheet1.xml", x =>
            {
                int i = x.IndexOf("<x:hyperlink ref=\"A1\"", StringComparison.Ordinal);
                i.Should().BeGreaterThan(0, "the hyperlink element must exist");
                int e = x.IndexOf("/>", i, StringComparison.Ordinal);
                return x[..i] + "<x:hyperlink ref=\"A1\"" + x[e..];
            }, required: true);

            using var opened = Workbook.Open(path);
            opened["S"]["A1"].GetHyperlink().Should().BeNull();
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Comment_AuthorId_OutOfRange_Fails_Loud()
    {
        // GetCommentAuthor on an authorId past the <authors> list throws
        // MalformedFileException (I-83). GetComment (the text) stays readable —
        // the corruption is confined to the author resolution.
        var path = AnnotatedFile();
        try
        {
            RewriteEntry(path, "xl/comments1.xml",
                x => x.Replace("authorId=\"0\"", "authorId=\"99\""), required: true);

            Action act = () => { using var wb = Workbook.Open(path); _ = wb["S"]["B2"].GetCommentAuthor(); };
            act.Should().Throw<MalformedFileException>();

            using var opened = Workbook.Open(path);
            opened["S"]["B2"].GetComment().Should().Be("note");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Comment_AuthorId_NonInteger_Fails_Loud()
    {
        // Was the one deliberate strictness divergence from NPOI (whose XML
        // deserializer silently coerced a non-integer @authorId to author "");
        // I-83 fail-loud is the contract.
        var path = AnnotatedFile();
        try
        {
            RewriteEntry(path, "xl/comments1.xml",
                x => x.Replace("authorId=\"0\"", "authorId=\"xyz\""), required: true);

            Action act = () => { using var wb = Workbook.Open(path); _ = wb["S"]["B2"].GetCommentAuthor(); };
            act.Should().Throw<MalformedFileException>();
        }
        finally { File.Delete(path); }
    }

    // ---- R-37: corrupted workbook-sheet relationship ids ----------------
    // Found by the deep-fuzz harness (a single bit flip in xl/workbook.xml
    // produced a raw NullReferenceException from Open). All three breach
    // shapes classify to MalformedFileException naming the sheet.

    private static string CorruptWorkbookXmlFile(Func<string, string> transform)
    {
        var path = TempPath();
        using (var wb = Workbook.Create())
        {
            wb.AddSheet("S")["A1"].SetString("hello");
            wb.Save(path);
        }
        RewriteEntry(path, "xl/workbook.xml", transform, required: true);
        return path;
    }

    [Fact]
    public void Sheet_Missing_RelId_Fails_Loud()
    {
        // Mangle the r:id attribute NAME so sheet.Id parses as absent.
        var path = CorruptWorkbookXmlFile(x => x.Replace("r:id=", "r:iX="));
        try
        {
            Action act = () => { using var wb = Workbook.Open(path); };
            act.Should().Throw<MalformedFileException>()
                .WithMessage("*missing its relationship id*");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Sheet_Dangling_RelId_Fails_Loud()
    {
        var path = CorruptWorkbookXmlFile(x =>
            System.Text.RegularExpressions.Regex.Replace(x, "r:id=\"[^\"]+\"", "r:id=\"rId999\""));
        try
        {
            Action act = () => { using var wb = Workbook.Open(path); };
            act.Should().Throw<MalformedFileException>()
                .WithMessage("*rId999*does not exist*");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Sheet_Targeting_NonWorksheet_Part_Fails_Loud()
    {
        // Point the sheet's r:id at the styles part — a genuinely non-sheet
        // target still fails loud. (Since I-92/R-38, ChartsheetPart and
        // DialogsheetPart targets open as placeholders instead — see
        // ChartsheetTests; only non-sheet parts like this land here.)
        var path = TempPath();
        using (var wb = Workbook.Create())
        {
            wb.AddSheet("S")["A1"].SetString("hello");
            wb.Save(path);
        }
        string? stylesId = null;
        RewriteEntry(path, "xl/_rels/workbook.xml.rels", x =>
        {
            var m = System.Text.RegularExpressions.Regex.Match(x, "<Relationship(?=[^>]*styles)[^>]*Id=\"([^\"]+)\"");
            m.Success.Should().BeTrue("the styles relationship must exist");
            stylesId = m.Groups[1].Value;
            return x + " ";   // satisfied the must-change guard; content-equivalent
        }, required: true);
        RewriteEntry(path, "xl/workbook.xml", x =>
            System.Text.RegularExpressions.Regex.Replace(x, "r:id=\"[^\"]+\"", $"r:id=\"{stylesId}\""), required: true);
        try
        {
            Action act = () => { using var wb = Workbook.Open(path); };
            act.Should().Throw<MalformedFileException>()
                .WithMessage("*not a sheet part*");
        }
        finally { File.Delete(path); }
    }
}
