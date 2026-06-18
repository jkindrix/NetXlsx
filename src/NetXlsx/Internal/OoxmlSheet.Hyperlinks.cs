// I-82 engine swap — Open XML SDK-backed hyperlinks (formulas/comments/
// hyperlinks slice).
//
// Mirrors the NPOI engine's ICell.Hyperlink / GetHyperlink contract (decision
// I13): the target scheme is sniffed — http(s)://, mailto:, file:// become
// EXTERNAL package relationships (TargetMode="External", referenced from the
// <hyperlink> by r:id); an internal "#Sheet!Range" form becomes @location with
// the '#' stripped and needs no relationship. Anything else is rejected.
//
// <hyperlinks> is a strict-sequence CT_Worksheet child (it slots between
// <dataValidations> and <printOptions>), so the container is placed by
// OoxmlSchemaOrder (SDK-quirk #8). External targets are stored VERBATIM — the
// packaging layer writes the Uri's OriginalString into the .rels Target and
// round-trips it, so GetHyperlink reports exactly what the caller wrote (no
// URI canonicalization; verified against the NPOI engine).

using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    // Applies (or replaces) the hyperlink on one cell. The caller (OoxmlCell)
    // owns the cell-text semantics; this owns the <hyperlink> element + the
    // package relationship.
    internal void SetHyperlink(int row, int col, string target, string? display)
    {
        var (isExternal, address) = SniffHyperlinkScheme(target);
        var reference = CellAddress.Format(row, col);

        var links = OoxmlSchemaOrder.GetOrInsert(Worksheet, static () => new S.Hyperlinks());

        // Replace semantics (NPOI parity): one hyperlink per cell. Drop the old
        // element and, when it carried an external relationship, the relationship
        // too (no orphaned .rels entry).
        var existing = links.Elements<S.Hyperlink>().FirstOrDefault(h => h.Reference?.Value == reference);
        if (existing is not null)
        {
            if (existing.Id?.Value is { } oldRelId)
            {
                var oldRel = _worksheetPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == oldRelId);
                if (oldRel is not null) _worksheetPart.DeleteReferenceRelationship(oldRel);
            }
            existing.Remove();
        }

        var link = new S.Hyperlink { Reference = reference };
        if (isExternal)
        {
            Uri uri;
            try
            {
                uri = new Uri(address, UriKind.Absolute);
            }
            catch (UriFormatException ex)
            {
                throw new ArgumentException(
                    $"hyperlink target '{target}' is not a valid URI: {ex.Message}", nameof(target), ex);
            }
            link.Id = _worksheetPart.AddHyperlinkRelationship(uri, isExternal: true).Id;
        }
        else
        {
            link.Location = address;
        }
        if (display is not null) link.Display = display;
        links.AppendChild(link);
    }

    // Removes the cell's <hyperlink> element (I-91 removal family). Reuses
    // SetHyperlink's replace-cleanup discipline: an external link's reference
    // relationship goes with the element (no orphaned .rels entry), and the
    // <hyperlinks> container is dropped once its last child leaves (SDK-quirk
    // #7 — Excel is friendlier with it absent). Idempotent: a no-op when the
    // cell carries no hyperlink. The cell's display text is left untouched
    // (R-10 — removal is the explicit way off; editing text keeps the link).
    internal void RemoveHyperlink(int row, int col)
    {
        var reference = CellAddress.Format(row, col);
        var links = Worksheet.GetFirstChild<S.Hyperlinks>();
        var link = links?.Elements<S.Hyperlink>().FirstOrDefault(h => h.Reference?.Value == reference);
        if (link is null) return;

        if (link.Id?.Value is { } relId)
        {
            var rel = _worksheetPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            if (rel is not null) _worksheetPart.DeleteReferenceRelationship(rel);
        }
        link.Remove();

        if (!links!.Elements<S.Hyperlink>().Any()) links.Remove();
    }

    // The cell's hyperlink target, or null when none is attached. External
    // links resolve their r:id to the package relationship and report the
    // verbatim target; internal links report @location (the '#'-stripped body,
    // NPOI parity).
    internal string? GetHyperlink(int row, int col)
    {
        var reference = CellAddress.Format(row, col);
        var link = Worksheet.GetFirstChild<S.Hyperlinks>()?.Elements<S.Hyperlink>()
            .FirstOrDefault(h => h.Reference?.Value == reference);
        if (link is null) return null;

        if (link.Id?.Value is { } relId)
        {
            // Fail loud on a dangling r:id (decision I-83): the reference is
            // genuine corruption — OOXML defines no fallback for it, and the
            // NPOI engine refuses such a file outright.
            var rel = _worksheetPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId)
                ?? throw new MalformedFileException(
                    $"cell {reference}: hyperlink r:id '{relId}' has no matching relationship");
            return rel.Uri.OriginalString;
        }
        // No r:id and no @location is a meaningless-but-legal degenerate
        // (both attributes are schema-optional); report "no hyperlink".
        return link.Location?.Value;
    }

    // Decision I13: scheme-sniff the target. The original text is preserved
    // verbatim; the external/internal split tells OPC how to interpret it.
    private static (bool IsExternal, string Address) SniffHyperlinkScheme(string target)
    {
        if (target.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || target.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            || target.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
            || target.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
            return (true, target);
        if (target.StartsWith('#'))
            return (false, target.Substring(1));
        throw new ArgumentException(
            $"hyperlink target '{target}' uses an unsupported scheme. " +
            "Supported: http(s)://, mailto:, file://, internal #Sheet!Range " +
            "(decision I13).", nameof(target));
    }
}
