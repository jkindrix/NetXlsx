// I-89 (S2 memo amendment, 2026-06-11) — the lazy default-theme discipline is
// STRUCTURAL, not conventional: every site that writes a theme-indexed color
// into output XML must route through the OoxmlWorkbook.EnsureThemePart choke
// point (directly, or via the style pool's OnThemeColorWrite hook). There is
// no natural "first theme write" location — the sites are scattered across
// the style pool, the drawing layer, and the rich-text path — so this test
// enumerates them from the SOURCE TREE and fails when a new theme-color
// write appears without the guard.
//
// What counts as a theme-color write (XML level):
//   - constructing a drawingML <a:schemeClr> (new A.SchemeColor / SchemeColor)
//   - setting CT_Color/@theme (Theme = (uint)..., Theme = <n>u)
// Model-level record members (FontColorTheme = ..., BackgroundTheme = ...)
// are NOT writes; the regexes below exclude them via a letter lookbehind.
//
// The one deliberate exemption (documented at both ends): the created
// stylesheet's scaffolding font 0 carries <color theme="1"/> in EVERY
// workbook — treating it as a trigger would make the lazy embed eager and
// break the byte-identity guarantee for theme-free workbooks.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests.Engine;

public class ThemeColorWriteSiteTests
{
    // file name -> the guard token its theme-color writes must route through.
    private static readonly Dictionary<string, string> ExpectedSites = new(StringComparer.Ordinal)
    {
        // Style pool: fill/font/border theme allocations funnel through
        // NoteThemeColorWrite (the DOM engine wires the hook to
        // EnsureThemePart; the streaming engine reads UsesThemeColors at
        // assembly). Also hosts the scaffolding font-0 exemption.
        ["OoxmlStylePool.cs"] = "NoteThemeColorWrite",
        // Drawing layer: schemeClr writes guard directly at the workbook.
        ["OoxmlPicture.cs"] = "EnsureThemePart",
        ["OoxmlSheet.Shapes.cs"] = "EnsureThemePart",
        ["OoxmlSheet.Charts.cs"] = "EnsureThemePart",
    };

    private static readonly Regex ThemeWriteRegex = new(
        // new A.SchemeColor / new SchemeColor — drawingML scheme references;
        // Theme = (uint)... / Theme = 1u — CT_Color theme-index assignments.
        // (?<![A-Za-z]) keeps BackgroundTheme/FontColorTheme/ColorTheme record
        // members (and ThemePart) out.
        @"new\s+(?:[A-Za-z]+\.)?SchemeColor\b|(?<![A-Za-z])Theme\s*=\s*(?:\(uint\)|\d+[uU])",
        RegexOptions.Compiled);

    [Fact]
    public void Every_Theme_Color_Write_Site_Is_Guarded()
    {
        var srcDir = LibrarySourceDir();
        var hits = new Dictionary<string, List<string>>(StringComparer.Ordinal);

        foreach (var file in Directory.EnumerateFiles(srcDir, "*.cs", SearchOption.AllDirectories))
        {
            if (file.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}")
                || file.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}"))
                continue;
            var text = File.ReadAllText(file);
            var matches = ThemeWriteRegex.Matches(text);
            if (matches.Count > 0)
                hits[Path.GetFileName(file)] = matches.Select(m => m.Value).Distinct().ToList();
        }

        // 1. No theme-color write outside the enumerated sites.
        var unexpected = hits.Keys.Except(ExpectedSites.Keys).ToList();
        unexpected.Should().BeEmpty(
            "every theme-color write site must call OoxmlWorkbook.EnsureThemePart (directly or via " +
            "the style pool's OnThemeColorWrite hook) and be enumerated in this test — found " +
            $"unguarded theme-color writes in: {string.Join(", ", unexpected.Select(f => $"{f} ({string.Join("; ", hits[f])})"))}");

        // 2. Every enumerated site still exists and still writes theme colors
        //    (a site that stops writing them should leave this list).
        var stale = ExpectedSites.Keys.Except(hits.Keys).ToList();
        stale.Should().BeEmpty("these files no longer contain theme-color writes; prune them from ExpectedSites");

        // 3. Each site carries its guard.
        foreach (var (file, guard) in ExpectedSites)
        {
            var path = Directory.EnumerateFiles(srcDir, file, SearchOption.AllDirectories).Single();
            File.ReadAllText(path).Should().Contain(guard,
                $"{file} writes theme-indexed colors and must route them through {guard}");
        }
    }

    [Fact]
    public void Every_Run_Properties_Builder_Caller_Is_Guarded()
    {
        // Rich-text run colors are built by the static
        // OoxmlStylePool.BuildRunProperties — outside both the pool hook and
        // the drawing guards — so every CALLER owns the EnsureThemePart call.
        // Today that is OoxmlCell.SetRichText; a future surface that grows
        // formatted runs (e.g. rich comments) lands here automatically.
        var srcDir = LibrarySourceDir();
        foreach (var file in Directory.EnumerateFiles(srcDir, "*.cs", SearchOption.AllDirectories))
        {
            var name = Path.GetFileName(file);
            if (name == "OoxmlStylePool.cs") continue; // the builder itself
            var text = File.ReadAllText(file);
            if (!text.Contains("BuildRunProperties(")) continue;
            text.Should().Contain("EnsureThemePart",
                $"{name} writes rich-text runs (BuildRunProperties) and must guard run ColorTheme " +
                "with EnsureThemePart before emission");
        }
    }

    private static string LibrarySourceDir([CallerFilePath] string thisFile = "")
    {
        // tests/NetXlsx.Tests/Engine/<this file> -> repo root -> src/NetXlsx.
        var repoRoot = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(thisFile)!, "..", "..", ".."));
        var srcDir = Path.Combine(repoRoot, "src", "NetXlsx");
        // Fail loud if the source tree is absent (same philosophy as the
        // fixture-provenance gate): a silently-skipped structural test rots.
        Directory.Exists(srcDir).Should().BeTrue(
            $"the structural write-site gate needs the library source tree at {srcDir}");
        return srcDir;
    }
}
