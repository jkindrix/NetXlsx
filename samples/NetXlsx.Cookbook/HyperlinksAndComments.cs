// Cookbook recipe 8 — HyperlinksAndComments
//
// Per docs/design.md §8.1: "Annotated cells with comments and external
// hyperlinks."
//
// Unblocked by v0.7 sub-slice C (ICell.Comment / Hyperlink). Comment
// authors default to "NetXlsx" per decision I11 — explicit
// attribution, no PII leak via Environment.UserName.

using System.Threading.Tasks;
using NetXlsx;

namespace NetXlsx.Cookbook.Recipes;

/// <summary>
/// Builds a small "release notes" sheet where each row points at a
/// reference URL (hyperlink) and carries a reviewer note (comment).
/// Demonstrates the four supported hyperlink schemes per decision I13:
/// <c>http(s)://</c>, <c>mailto:</c>, <c>file://</c>, and the internal
/// <c>#Sheet!Range</c> form.
/// </summary>
public static class HyperlinksAndComments
{
    /// <summary>Output sheet name.</summary>
    public const string SheetName = "ReleaseNotes";

    /// <summary>Runs the recipe.</summary>
    public static async Task Run(string outputPath)
    {
        using var wb = Workbook.Create();

        // A second sheet so the internal #Sheet!Range link has somewhere
        // to point.
        var changelog = wb.AddSheet("Changelog");
        changelog.AppendRow().Set(1, "Version").Set(2, "Notes");
        changelog.AppendRow().Set(1, "1.0.0").Set(2, "Initial release");

        var notes = wb.AddSheet(SheetName);
        notes.AppendRow().Set(1, "Item").Set(2, "Link");

        // 1. External https link with an explicit display string.
        notes.AppendRow().Set(1, "Public site");
        notes["B3"].Hyperlink("https://example.com/releases", display: "Release page");

        // 2. mailto: link — no display, falls back to the raw target.
        notes.AppendRow().Set(1, "Maintainer");
        notes["B4"].Hyperlink("mailto:maintainer@example.com");

        // 3. file:// link to a sibling artifact on a network share.
        notes.AppendRow().Set(1, "Release artifact");
        notes["B5"].Hyperlink("file:///net/share/releases/1.0.0/notes.pdf",
                              display: "1.0.0/notes.pdf");

        // 4. Internal cross-sheet link.
        notes.AppendRow().Set(1, "See also");
        notes["B6"].Hyperlink("#Changelog!A2", display: "Changelog row");

        // Comments demonstrate the default author (I11) and an explicit
        // author override.
        notes["A3"].Comment("Always link the canonical release page.");
        notes["A4"].Comment("Distribution list, not a single human.",
                            author: "release-bot");
        notes["A5"].Comment("Path is canonical on the build server.");

        notes.FreezeRows(1);

        await wb.SaveAsync(outputPath);
    }
}
