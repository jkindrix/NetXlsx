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
// container's known child-type sequence and placing the new element before the
// first existing sibling that must follow it. Every structural slice from here
// (panes, grouping, protection, drawings, conditional formatting, validation,
// tables) hits the same ordered-container problem, so the fix lives here once.
//
// The sequences below are the ECMA-376 CT_Worksheet / CT_Workbook child orders;
// the SDK element types were confirmed by reflection against
// DocumentFormat.OpenXml 3.5.1. A handful of rare legacy elements (smartTags,
// smartTagPr, smartTagTypes) are not modeled as distinct elements by this SDK
// version and are omitted; an element whose type is not listed is treated as
// sorting last (see RankOf), so a round-tripped unmodeled sibling is never
// displaced ahead of a known element.

using System;
using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal static class OoxmlSchemaOrder
{
    // CT_Worksheet child sequence (ECMA-376 §18.3.1.99).
    private static readonly Type[] WorksheetOrder =
    {
        typeof(S.SheetProperties),            // sheetPr
        typeof(S.SheetDimension),             // dimension
        typeof(S.SheetViews),                 // sheetViews
        typeof(S.SheetFormatProperties),      // sheetFormatPr
        typeof(S.Columns),                    // cols
        typeof(S.SheetData),                  // sheetData
        typeof(S.SheetCalculationProperties), // sheetCalcPr
        typeof(S.SheetProtection),            // sheetProtection
        typeof(S.ProtectedRanges),            // protectedRanges
        typeof(S.Scenarios),                  // scenarios
        typeof(S.AutoFilter),                 // autoFilter
        typeof(S.SortState),                  // sortState
        typeof(S.DataConsolidate),            // dataConsolidate
        typeof(S.CustomSheetViews),           // customSheetViews
        typeof(S.MergeCells),                 // mergeCells
        typeof(S.PhoneticProperties),         // phoneticPr
        typeof(S.ConditionalFormatting),      // conditionalFormatting (0..*)
        typeof(S.DataValidations),            // dataValidations
        typeof(S.Hyperlinks),                 // hyperlinks
        typeof(S.PrintOptions),               // printOptions
        typeof(S.PageMargins),                // pageMargins
        typeof(S.PageSetup),                  // pageSetup
        typeof(S.HeaderFooter),               // headerFooter
        typeof(S.RowBreaks),                  // rowBreaks
        typeof(S.ColumnBreaks),               // colBreaks
        typeof(S.CustomProperties),           // customProperties
        typeof(S.CellWatches),                // cellWatches
        typeof(S.IgnoredErrors),              // ignoredErrors
        typeof(S.Drawing),                    // drawing
        typeof(S.DrawingHeaderFooter),        // drawingHF
        typeof(S.Picture),                    // picture
        typeof(S.OleObjects),                 // oleObjects
        typeof(S.Controls),                   // controls
        typeof(S.WebPublishItems),            // webPublishItems
        typeof(S.TableParts),                 // tableParts
        typeof(S.WorksheetExtensionList),     // extLst
    };

    // CT_Workbook child sequence (ECMA-376 §18.2.27).
    private static readonly Type[] WorkbookOrder =
    {
        typeof(S.FileVersion),                // fileVersion
        typeof(S.FileSharing),                // fileSharing
        typeof(S.WorkbookProperties),         // workbookPr
        typeof(S.WorkbookProtection),         // workbookProtection
        typeof(S.BookViews),                  // bookViews
        typeof(S.Sheets),                     // sheets
        typeof(S.FunctionGroups),             // functionGroups
        typeof(S.ExternalReferences),         // externalReferences
        typeof(S.DefinedNames),               // definedNames
        typeof(S.CalculationProperties),      // calcPr
        typeof(S.OleSize),                    // oleSize
        typeof(S.CustomWorkbookViews),        // customWorkbookViews
        typeof(S.PivotCaches),                // pivotCaches
        typeof(S.WebPublishing),              // webPublishing
        typeof(S.FileRecoveryProperties),     // fileRecoveryPr
        typeof(S.WebPublishObjects),          // webPublishObjects
        typeof(S.WorkbookExtensionList),      // extLst
    };

    /// <summary>
    /// Inserts <paramref name="child"/> into <paramref name="worksheet"/> at its
    /// correct position in the CT_Worksheet child sequence, or returns the existing
    /// child of the same type if one is already present.
    /// </summary>
    internal static T GetOrInsert<T>(S.Worksheet worksheet, Func<T> factory)
        where T : OpenXmlElement
    {
        var existing = worksheet.GetFirstChild<T>();
        if (existing is not null) return existing;
        return InsertOrdered(worksheet, factory(), WorksheetOrder);
    }

    /// <summary>
    /// Inserts <paramref name="child"/> into <paramref name="workbook"/> at its
    /// correct position in the CT_Workbook child sequence, or returns the existing
    /// child of the same type if one is already present.
    /// </summary>
    internal static T GetOrInsert<T>(S.Workbook workbook, Func<T> factory)
        where T : OpenXmlElement
    {
        var existing = workbook.GetFirstChild<T>();
        if (existing is not null) return existing;
        return InsertOrdered(workbook, factory(), WorkbookOrder);
    }

    private static T InsertOrdered<T>(OpenXmlElement parent, T child, Type[] order)
        where T : OpenXmlElement
    {
        int childRank = RankOf(child.GetType(), order);
        foreach (var sibling in parent.ChildElements)
        {
            // Place the new child before the first existing sibling that must
            // follow it. A sibling of equal rank (same repeatable type) does not
            // displace us — we append after it.
            if (RankOf(sibling.GetType(), order) > childRank)
            {
                parent.InsertBefore(child, sibling);
                return child;
            }
        }
        parent.AppendChild(child);
        return child;
    }

    // Position of a child type in the container's sequence. An unlisted type sorts
    // last (int.MaxValue) so a round-tripped unmodeled element is never pushed
    // ahead of a known element we are inserting.
    private static int RankOf(Type type, Type[] order)
    {
        for (int i = 0; i < order.Length; i++)
            if (order[i] == type) return i;
        return int.MaxValue;
    }
}
