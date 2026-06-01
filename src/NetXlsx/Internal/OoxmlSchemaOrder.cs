// I-82 engine swap — schema-ordered child insertion for ordered OOXML containers.
//
// CT_Worksheet and CT_Workbook are strict-sequence complex types: their children
// must appear in a fixed order, and the SDK's AppendChild / InsertAfter(anchor) do
// NOT reorder (SDK-quirk #3). The 5a structure code inserted <mergeCells> right
// after <sheetData> and <definedNames> right after <sheets>. That is correct for
// workbooks the engine *creates* (those have no intervening siblings), but on an
// *opened* file that already carries a legal sibling that must sit between the
// anchor and the new element — e.g. <autoFilter> (ubiquitous in real files) sits
// between <sheetData> and <mergeCells>, and <functionGroups>/<externalReferences>
// sit between <sheets> and <definedNames> — the bare InsertAfter(anchor) emits
// out-of-order XML, which OpenXmlValidator rejects with
// Sch_UnexpectedElementContentExpectingComplex and which would be a regression vs
// the NPOI engine at cutover (NPOI orders internally).
//
// This helper inserts a child at its correct schema position by walking the
// container's child sequence and placing the new element before the first existing
// sibling that must follow it. Every structural slice from here (panes, grouping,
// protection, drawings, conditional formatting, validation, tables) hits the same
// ordered-container problem, so the fix lives here once.
//
// SOURCE OF TRUTH FOR THE ORDER LISTS. The two arrays below are the *local element
// names* of each container's child sequence, taken verbatim from
// DocumentFormat.OpenXml 3.5.1's own compiled schema particle — i.e. the exact
// order OpenXmlValidator enforces — not hand-transcribed from ECMA-376.
// SchemaOrderCanonicalTests reflects that same SDK particle metadata and fails the
// build if either array drifts (a missing, extra, or misordered name), so an
// incomplete list cannot ship. This guard exists because the previous hand-derived
// Type[] form silently omitted three elements that no created-from-scratch fixture
// could catch: <legacyDrawing>/<legacyDrawingHF> (worksheet) and — caught only by
// the machine-check — <absPath> (workbook).
//
// WHY KEY BY LOCAL NAME, NOT CLR Type. The schema sequence these lists encode is
// defined over element *qualified names*, and the machine-check derives its truth
// from the SDK particle as names — so a name-keyed list maps to that truth 1:1, with
// no Type<->name round-trip to get wrong. It is also more robust: the workbook
// sequence includes <x15ac:absPath> at ordinal 3 (Excel emits it routinely), whose
// element class hides in an Office2013 extension namespace
// (DocumentFormat.OpenXml.Office2013.ExcelAc.AbsolutePath) — exactly the kind of
// awkward-to-name type a Type[] list invites omitting. Keying on
// OpenXmlElement.LocalName ranks such an element correctly whether an opened file
// round-trips it as its typed class or as an OpenXmlUnknownElement. Local names are
// unique within each of these two sequences (the lone non-main-namespace member,
// absPath, shares its name with nothing), so name keys do not collide. The lost
// compile-time safety of typeof(...) is recovered by SchemaOrderCanonicalTests,
// which fails the build on any name that does not match the SDK particle.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal static class OoxmlSchemaOrder
{
    // CT_Worksheet child sequence (ECMA-376 §18.3.1.99), verbatim from the SDK's
    // compiled particle. Verified element-for-element by SchemaOrderCanonicalTests.
    internal static readonly string[] WorksheetChildOrder =
    {
        "sheetPr",
        "dimension",
        "sheetViews",
        "sheetFormatPr",
        "cols",
        "sheetData",
        "sheetCalcPr",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "autoFilter",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
        "mergeCells",
        "phoneticPr",
        "conditionalFormatting",  // 0..*
        "dataValidations",
        "hyperlinks",
        "printOptions",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "rowBreaks",
        "colBreaks",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "drawing",
        "legacyDrawing",          // sits between <drawing> and <drawingHF>/<picture>
        "legacyDrawingHF",
        "drawingHF",
        "picture",
        "oleObjects",
        "controls",
        "webPublishItems",
        "tableParts",
        "extLst",
    };

    // CT_Workbook child sequence (ECMA-376 §18.2.27), verbatim from the SDK's
    // compiled particle. Verified element-for-element by SchemaOrderCanonicalTests.
    internal static readonly string[] WorkbookChildOrder =
    {
        "fileVersion",
        "fileSharing",
        "workbookPr",
        "absPath",                // x15ac:absPath (Office2013.ExcelAc.AbsolutePath) —
                                  // see header note on name-keying
        "workbookProtection",
        "bookViews",
        "sheets",
        "functionGroups",
        "externalReferences",
        "definedNames",
        "calcPr",
        "oleSize",
        "customWorkbookViews",
        "pivotCaches",
        "webPublishing",
        "fileRecoveryPr",
        "webPublishObjects",
        "extLst",
    };

    private static readonly Dictionary<string, int> WorksheetRank = BuildRank(WorksheetChildOrder);
    private static readonly Dictionary<string, int> WorkbookRank = BuildRank(WorkbookChildOrder);

    private static Dictionary<string, int> BuildRank(string[] order)
    {
        var rank = new Dictionary<string, int>(order.Length, StringComparer.Ordinal);
        for (int i = 0; i < order.Length; i++) rank[order[i]] = i;
        return rank;
    }

    /// <summary>
    /// Inserts the element produced by <paramref name="factory"/> into
    /// <paramref name="worksheet"/> at its correct position in the CT_Worksheet child
    /// sequence, or returns the existing child of the same type if one is present.
    /// </summary>
    internal static T GetOrInsert<T>(S.Worksheet worksheet, Func<T> factory)
        where T : OpenXmlElement
    {
        var existing = worksheet.GetFirstChild<T>();
        if (existing is not null) return existing;
        return InsertOrdered(worksheet, factory(), WorksheetRank);
    }

    /// <summary>
    /// Inserts the element produced by <paramref name="factory"/> into
    /// <paramref name="workbook"/> at its correct position in the CT_Workbook child
    /// sequence, or returns the existing child of the same type if one is present.
    /// </summary>
    internal static T GetOrInsert<T>(S.Workbook workbook, Func<T> factory)
        where T : OpenXmlElement
    {
        var existing = workbook.GetFirstChild<T>();
        if (existing is not null) return existing;
        return InsertOrdered(workbook, factory(), WorkbookRank);
    }

    private static T InsertOrdered<T>(OpenXmlElement parent, T child, Dictionary<string, int> rank)
        where T : OpenXmlElement
    {
        int childRank = RankOf(child, rank);
        foreach (var sibling in parent.ChildElements)
        {
            // Place the new child before the first existing sibling that must
            // follow it. A sibling of equal rank (same repeatable type) does not
            // displace us — we append after it.
            if (RankOf(sibling, rank) > childRank)
            {
                parent.InsertBefore(child, sibling);
                return child;
            }
        }
        parent.AppendChild(child);
        return child;
    }

    // Position of an element in the container's sequence, keyed by local name. An
    // element whose local name is not in the sequence sorts last (int.MaxValue) so a
    // round-tripped unmodeled element is never pushed ahead of a known element we are
    // inserting.
    private static int RankOf(OpenXmlElement element, Dictionary<string, int> rank)
        => rank.TryGetValue(element.LocalName, out int r) ? r : int.MaxValue;
}
