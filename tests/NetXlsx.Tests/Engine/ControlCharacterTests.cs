// I-88 (ledger R-3): lossless ST_Xstring escaping for user content,
// fail-fast for name/formula surfaces. Pre-fix, XML-invalid characters in
// SetString/rich text/Comment passed the setter silently and exploded at
// Save as a raw SDK ArgumentException — and CR (0x0D) was SILENTLY
// DESTROYED ("C\rD" read back "C<LF>D"). Closes R-30's control-char /
// lone-surrogate test gap as a side effect.

using System;
using System.Collections;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using NetXlsx.Tests.SourceGen;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class ControlCharacterTests
{
    private static IWorkbook RoundTrip(IWorkbook wb)
    {
        var ms = new MemoryStream();
        wb.Save(ms);
        wb.Dispose();
        ms.Position = 0;
        return Workbook.Open(ms);
    }

    // ---- The full 0x00–0x1F partition ----------------------------------

    public static TheoryData<int> EscapedControlChars()
    {
        var data = new TheoryData<int>();
        for (int c = 0; c <= 0x1F; c++)
            if (c != 0x09 && c != 0x0A) data.Add(c);   // tab + LF pass through
        return data;
    }

    [Theory]
    [MemberData(nameof(EscapedControlChars))]
    public void Control_Char_Round_Trips_Through_SetString(int code)
    {
        var text = $"a{(char)code}b";
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetString(text);

        using var read = RoundTrip(wb);
        read["S"]["A1"].GetString().Should().Be(text,
            $"U+{code:X4} must round-trip losslessly via the _xHHHH_ escape");
    }

    [Theory]
    [InlineData("\t")]
    [InlineData("\n")]
    public void Tab_And_LF_Pass_Through_Unescaped(string ch)
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetString($"a{ch}b");

        // The raw part carries the literal char, not an escape (MS-OI29500:
        // Office does not accept these escaped in element content).
        var xml = SavedOoxml.SheetXml(wb).ToString();
        xml.Should().NotContain("_x0009_").And.NotContain("_x000A_");

        using var read = RoundTrip(wb);
        read["S"]["A1"].GetString().Should().Be($"a{ch}b");
    }

    [Fact]
    public void CR_Is_No_Longer_Silently_Destroyed()
    {
        // The pre-I-88 behavior, repro-proven 2026-06-11: "C\rD" → "C\nD"
        // (XmlWriter newline normalization + XML 1.0 §2.11). Now escaped.
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        s["A1"].SetString("C\rD");
        s["A2"].SetString("C\r\nD");

        using var read = RoundTrip(wb);
        read["S"]["A1"].GetString().Should().Be("C\rD");
        read["S"]["A2"].GetString().Should().Be("C\r\nD",
            "CRLF must not collapse to a lone LF");
    }

    [Theory]
    [InlineData("\uFFFE")]
    [InlineData("\uFFFF")]
    [InlineData("\uD800")]   // lone high surrogate
    [InlineData("\uDC00")]   // lone low surrogate
    public void Noncharacters_And_Lone_Surrogates_Round_Trip(string ch)
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetString($"a{ch}b");

        using var read = RoundTrip(wb);
        read["S"]["A1"].GetString().Should().Be($"a{ch}b");
    }

    [Fact]
    public void Paired_Surrogates_Still_Pass_Through_Unescaped()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetString("rocket 🚀 ship");

        SavedOoxml.SheetXml(wb).ToString().Should().NotContain("_xD83D_",
            "a well-formed surrogate pair is ordinary text");
        using var read = RoundTrip(wb);
        read["S"]["A1"].GetString().Should().Be("rocket 🚀 ship");
    }

    // ---- Literal-escape protection (_x005F_) ---------------------------

    [Fact]
    public void Literal_Escape_Pattern_In_User_Text_Round_Trips()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetString("file_x0008_name");

        var xml = SavedOoxml.SheetXml(wb).ToString();
        xml.Should().Contain("_x005F_x0008_", "the convention protects the literal's underscore");

        using var read = RoundTrip(wb);
        read["S"]["A1"].GetString().Should().Be("file_x0008_name");
    }

    [Fact]
    public void Emitted_Escape_Is_Uppercase_Hex()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetString("a\u001Fb");
        SavedOoxml.SheetXml(wb).ToString().Should().Contain("_x001F_",
            "Excel emits uppercase hex; so do we");
    }

    // ---- Decode of foreign-authored escapes ----------------------------

    [Theory]
    [InlineData("_x0008_", "\b")]            // Excel's own shape
    [InlineData("_x000d_", "\r")]            // decode is hex-case-insensitive
    [InlineData("_X0041_", "_X0041_")]       // uppercase X is NOT an escape (ClosedXML pins this)
    [InlineData("_x00G1_", "_x00G1_")]       // non-hex digit — literal
    [InlineData("_x005F_x0041_", "_x0041_")] // protected literal decodes once
    public void Foreign_Authored_Escapes_Decode_Per_Convention(string stored, string expected)
    {
        // Inject the stored text verbatim into the part (the crafted-cell
        // pattern) — this is what an Excel-authored file carries.
        var path = Path.Combine(Path.GetTempPath(), $"netxlsx-i88-{Guid.NewGuid():N}.xlsx");
        using (var wb = Workbook.Create())
        {
            wb.AddSheet("S")["A1"].SetString("placeholder");
            wb.Save(path);
        }
        try
        {
            using (var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update))
            {
                var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
                string xml;
                using (var r = new StreamReader(entry.Open())) xml = r.ReadToEnd();
                xml = xml.Replace("placeholder", stored);
                entry.Delete();
                using var w = new StreamWriter(zip.CreateEntry("xl/worksheets/sheet1.xml").Open());
                w.Write(xml);
            }
            using var read = Workbook.Open(path);
            read["S"]["A1"].GetString().Should().Be(expected);
        }
        finally { File.Delete(path); }
    }

    // ---- Other content surfaces ----------------------------------------

    [Fact]
    public void Rich_Text_Runs_Round_Trip_Control_Chars()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetRichText(new RichText(
            new RichTextRun("bell\u0007") { Style = new RichTextStyle { Bold = true } },
            new RichTextRun("cr\r") { Style = new RichTextStyle { Italic = true } }));

        using var read = RoundTrip(wb);
        var rt = read["S"]["A1"].GetRichText();
        rt.Should().NotBeNull();
        rt!.Runs[0].Text.Should().Be("bell\u0007");
        rt.Runs[1].Text.Should().Be("cr\r");
        read["S"]["A1"].GetString().Should().Be("bell\u0007cr\r");
    }

    [Fact]
    public void Comments_Round_Trip_Control_Chars()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].Comment("line1\rline2\u0001", author: "QA");

        using var read = RoundTrip(wb);
        read["S"]["A1"].GetComment().Should().Be("line1\rline2\u0001");
    }

    [Fact]
    public void Streaming_SetString_Escapes_Identically()
    {
        using var ms = new MemoryStream();
        using (var wb = Workbook.CreateStreaming())
        {
            wb.AddSheet("Big").AppendRow().Set(1, "a\u0003b\rc");
            wb.Save(ms, leaveOpen: true);
        }
        ms.Position = 0;
        using var read = Workbook.Open(ms);
        read["Big"]["A1"].GetString().Should().Be("a\u0003b\rc");
    }

    [Fact]
    public void Generated_ReadRows_Round_Trips_Escaped_Content()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial record Item([property: Column(""Name"")] string Name);";
        var output = GeneratorHarness.RunWithFullReferences(src, out var compilation);
        output.CompilationDiagnostics.Should().NotContain(
            d => d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Error);
        var asm = GeneratorHarness.EmitAndLoad(compilation);
        var ext = asm.GetType("T.Item_SheetExtensions")!;
        var itemType = asm.GetType("T.Item")!;

        var items = Array.CreateInstance(itemType, 1);
        items.SetValue(Activator.CreateInstance(itemType, "bell\u0007cr\r"), 0);

        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var sheet = wb.AddSheet("I");
            sheet.AppendRow().Set(1, "Name");
            ext.GetMethod("AddRows")!.Invoke(null, new object[] { sheet, items });
            wb.Save(ms);
        }
        ms.Position = 0;
        using (var wb = Workbook.Open(ms))
        {
            var rows = (IEnumerable)ext.GetMethod("ReadRows")!.Invoke(null, new object?[] { wb["I"], 1 })!;
            var back = rows.Cast<object>().Single();
            itemType.GetProperty("Name")!.GetValue(back).Should().Be("bell\u0007cr\r");
        }
    }

    // ---- Fail-fast surfaces (half (a)) ----------------------------------

    [Fact]
    public void SetFormula_Rejects_XmlInvalid_Chars_At_The_Setter()
    {
        using var wb = Workbook.Create();
        var s = wb.AddSheet("S");
        Action act = () => s["A1"].SetFormula("=A1\u0001+1");
        act.Should().Throw<FormulaException>().WithMessage("*U+0001*");
    }

    [Fact]
    public void AddNamedRange_Rejects_XmlInvalid_Formula_Chars()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("S");
        Action act = () => wb.AddNamedRange("N", "S!$A$1\u0002");
        act.Should().Throw<FormulaException>().WithMessage("*U+0002*");
    }

    [Theory]
    [InlineData("bad\u0001name")]
    [InlineData("tab\tname")]   // XML-legal, but attribute normalization mutates it
    [InlineData("lf\nname")]
    public void AddSheet_Rejects_Control_Chars_In_Names(string name)
    {
        using var wb = Workbook.Create();
        Action act = () => wb.AddSheet(name);
        act.Should().Throw<SheetNameException>();
        Workbook.IsValidSheetName(name).Should().BeFalse();
        Workbook.IsValidSheetName(Workbook.SanitizeSheetName(name)).Should().BeTrue();
    }

    [Fact]
    public void Save_No_Longer_Throws_Raw_ArgumentException_For_Content()
    {
        // The original R-3 failure shape: control char in, raw save-time
        // ArgumentException out. Now the save just works.
        using var wb = Workbook.Create();
        wb.AddSheet("S")["A1"].SetString("\u0000\u0001\u001F");
        using var ms = new MemoryStream();
        Action act = () => wb.Save(ms);
        act.Should().NotThrow();
    }
}
