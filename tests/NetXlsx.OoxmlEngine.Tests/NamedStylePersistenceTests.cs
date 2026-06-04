// I-82 engine swap — named-style OOXML persistence (closeout slice).
//
// Closes the styles-slice deferral: RegisterStyle on the SDK engine now
// persists each name as a cellStyleXfs <xf> + <cellStyle> entry (the NPOI
// engine's I-67 round-trip), and a freshly opened workbook rehydrates its
// registry from the persisted entries. Mirrors the round-trip half of the
// NPOI-engine NamedStyleTests; the in-memory registry behavior (register/
// get/replace/case-insensitivity within a session) is covered by
// CellStyleTests since the styles slice.
//
// Documented divergence from the NPOI witness (oracle-dumped 2026-06-03):
// NPOI stamps builtinId="0" on every user <cellStyle> entry, but builtinId 0
// claims the *Normal builtin* per ECMA-376 — this engine writes no builtinId
// on user styles (Excel's own files don't either). Publicly invisible; both
// engines rehydrate each other's files identically.

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.OoxmlEngine.Tests;

public class NamedStylePersistenceTests
{
    private static readonly string[] s_threeNames = new[] { "Header", "Body", "Accent" };
    private static readonly string[] s_headerOnly = new[] { "Header" };

    private static string TempPath(string tag)
        => Path.Combine(Path.GetTempPath(), $"ooxml-named-{tag}-{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void Registered_Names_Survive_File_Roundtrip()
    {
        var path = TempPath("rt");
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                wb.RegisterStyle("Header", new CellStyle { Bold = true, FontSize = 14 });
                wb.RegisterStyle("Body", new CellStyle { Italic = true });
                wb.RegisterStyle("Accent", new CellStyle { Bold = true, FontColor = Color.Red });
                wb.AddSheet("S");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                wb.RegisteredStyleNames.Should().BeEquivalentTo(s_threeNames);
                var header = wb.GetRegisteredStyle("Header");
                header.Should().NotBeNull();
                header!.Bold.Should().Be(true);
                header.FontSize.Should().Be(14);
                wb.GetRegisteredStyle("Accent")!.FontColor.Should().Be(Color.Red);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Roundtrip_Preserves_Case_Insensitive_Lookup()
    {
        var path = TempPath("case");
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                wb.RegisterStyle("MyHeader", new CellStyle { Bold = true });
                wb.AddSheet("S");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                wb.GetRegisteredStyle("myheader").Should().NotBeNull();
                wb.GetRegisteredStyle("MYHEADER").Should().NotBeNull();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Workbook_Without_Named_Styles_Reads_Back_Empty()
    {
        var path = TempPath("empty");
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                wb.AddSheet("S")["A1"].SetString("hello");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                // The built-in "Normal" entry is not a user-registered name.
                wb.RegisteredStyleNames.Should().BeEmpty();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void RegisterStyle_Replacement_Updates_Existing_OOXML_Entry()
    {
        var path = TempPath("replace");
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                wb.RegisterStyle("X", new CellStyle { Bold = true });
                wb.RegisterStyle("X", new CellStyle { Italic = true });
                wb.AddSheet("S");
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                var s = wb.GetRegisteredStyle("X");
                s.Should().NotBeNull();
                s!.Italic.Should().Be(true);
                s.Bold.Should().BeNull("the second registration replaced the first");
                wb.RegisteredStyleNames.Should().ContainSingle();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void User_Entries_Carry_No_BuiltinId()
    {
        // builtinId 0 claims the Normal builtin (ECMA-376 §18.8.7) — NPOI
        // stamps it on user styles anyway; this engine deliberately does not.
        using var wb = Workbook.CreateOoxml();
        wb.RegisterStyle("Header", new CellStyle { Bold = true });
        wb.AddSheet("S");

        var ss = wb.Underlying.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        var entries = ss.GetFirstChild<S.CellStyles>()!.Elements<S.CellStyle>().ToList();
        entries.Should().HaveCount(2);
        entries.Single(e => e.Name!.Value == "Normal").BuiltinId!.Value.Should().Be(0u);
        entries.Single(e => e.Name!.Value == "Header").BuiltinId.Should().BeNull();
    }

    [Fact]
    public void Registered_Style_Mirrors_The_CellXfs_Component_Ids()
    {
        using var wb = Workbook.CreateOoxml();
        wb.RegisterStyle("Header", new CellStyle { Bold = true, FontSize = 14 });
        wb.AddSheet("S");

        var ss = wb.Underlying.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        var styleXfs = ss.GetFirstChild<S.CellStyleFormats>()!.Elements<S.CellFormat>().ToList();
        styleXfs.Should().HaveCount(2, "the Normal master xf plus the named entry");
        var named = styleXfs[1];
        named.FontId!.Value.Should().BeGreaterThan(0u);
        named.ApplyFont!.Value.Should().BeTrue();

        var entry = ss.GetFirstChild<S.CellStyles>()!.Elements<S.CellStyle>()
            .Single(e => e.Name!.Value == "Header");
        entry.FormatId!.Value.Should().Be(1u);
    }

    [Fact]
    public void Rehydration_Tolerates_The_Npoi_BuiltinId_Artifact()
    {
        // A file whose named styles were persisted by NPOI's I-67 path carries
        // builtinId="0" on USER entries (an NPOI witness artifact — it claims
        // the Normal builtin per ECMA-376). Rehydration must read such entries
        // anyway. The artifact is synthesized onto an engine-written file
        // (the NPOI engine that used to author this fixture is retired).
        var path = TempPath("npoi-artifact");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.RegisterStyle("Header", new CellStyle { Bold = true, FontSize = 14 });
                wb.AddSheet("S");
                var entry = wb.Underlying.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
                    .GetFirstChild<S.CellStyles>()!.Elements<S.CellStyle>()
                    .Single(e => e.Name!.Value == "Header");
                entry.BuiltinId = 0; // the NPOI artifact
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb.RegisteredStyleNames.Should().BeEquivalentTo(s_headerOnly);
                var header = wb.GetRegisteredStyle("Header");
                header!.Bold.Should().Be(true);
                header.FontSize.Should().Be(14);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Registering_Into_A_Stylesheet_Missing_The_Style_Tables_Seeds_Normal_First()
    {
        // Open-mutate blind-spot fixture (the stylesheet flavor of quirk #8's
        // discipline): a hand-stripped file whose stylesheet has NO
        // cellStyleXfs and NO cellStyles. RegisterStyle must (a) create both
        // tables in CT_Stylesheet order, (b) seed index 0 with the Normal
        // master so existing cellXfs xfId="0" references keep pointing at a
        // Normal-shaped parent, and (c) validate clean.
        var path = TempPath("stripped");
        try
        {
            using (var wb = Workbook.CreateOoxml())
            {
                var c = wb.AddSheet("S")["A1"];
                c.SetString("x");
                c.Style(new CellStyle { Bold = true });
                wb.Save(path);
            }
            StripStyleTables(path);

            using (var wb = Workbook.OpenOoxml(path))
            {
                wb.RegisteredStyleNames.Should().BeEmpty("the tables were stripped");
                wb.RegisterStyle("Header", new CellStyle { Bold = true, FontSize = 14 });

                var ss = wb.Underlying.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                var styleXfs = ss.GetFirstChild<S.CellStyleFormats>()!.Elements<S.CellFormat>().ToList();
                styleXfs.Should().HaveCount(2, "Normal master seeded at 0, named entry at 1");
                styleXfs[0].FontId!.Value.Should().Be(0u);

                var names = ss.GetFirstChild<S.CellStyles>()!.Elements<S.CellStyle>()
                    .Select(e => e.Name!.Value).ToList();
                names.Should().Equal("Normal", "Header");

                OpenXmlValidationGate.AssertValid(wb);
                wb.Save(path);
            }
            using (var wb = Workbook.OpenOoxml(path))
            {
                wb.RegisteredStyleNames.Should().BeEquivalentTo(s_headerOnly);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    // Removes the <cellStyleXfs> and <cellStyles> elements from xl/styles.xml —
    // legal per the schema (both are optional), and the shape minimal writers
    // produce.
    private static void StripStyleTables(string path)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = zip.GetEntry("xl/styles.xml")!;
        string xml;
        using (var reader = new StreamReader(entry.Open()))
            xml = reader.ReadToEnd();

        xml = System.Text.RegularExpressions.Regex.Replace(
            xml, "<cellStyleXfs[^>]*>.*?</cellStyleXfs>|<cellStyleXfs[^>]*/>", "",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        xml = System.Text.RegularExpressions.Regex.Replace(
            xml, "<cellStyles[^>]*>.*?</cellStyles>|<cellStyles[^>]*/>", "",
            System.Text.RegularExpressions.RegexOptions.Singleline);

        entry.Delete();
        var fresh = zip.CreateEntry("xl/styles.xml");
        using var writer = new StreamWriter(fresh.Open());
        writer.Write(xml);
    }
}
