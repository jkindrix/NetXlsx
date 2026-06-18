// Internal abstraction over the two sheet shapes the engine holds in its
// sheet collection (decision I-92): the grid-backed OoxmlSheet and the
// non-grid OoxmlChartsheet placeholder (chartsheet / dialogsheet). The
// workbook's lifecycle code (rename / move / remove, the open-time index,
// the visible-sheet guard) operates through these members without caring
// which concrete shape backs a given <sheet> entry.

using DocumentFormat.OpenXml.Packaging;

namespace NetXlsx;

internal interface IOoxmlSheet : ISheet
{
    /// <summary>The owning workbook (non-throwing — used by ownership checks).</summary>
    OoxmlWorkbook WorkbookInternal { get; }

    /// <summary>
    /// The OPC part backing this sheet's <c>&lt;sheet&gt;</c> entry — a
    /// <see cref="WorksheetPart"/>, <see cref="ChartsheetPart"/>, or
    /// <see cref="DialogsheetPart"/>. The workbook resolves the
    /// <c>&lt;sheet&gt;</c> element and tears the part down through this.
    /// </summary>
    OpenXmlPart SheetPartInternal { get; }

    /// <summary>
    /// Non-throwing visibility read for the workbook-side last-visible-sheet
    /// guard (must not trip the disposed/removed guards).
    /// </summary>
    bool IsHiddenInternal { get; }

    /// <summary>Keeps the wrapper's cached name in sync on rename.</summary>
    void SetNameInternal(string name);

    /// <summary>Marks the wrapper a tombstone after removal (one-way).</summary>
    void MarkRemoved();

    /// <summary>
    /// Rewrites this sheet's own formula-shaped surfaces from
    /// <paramref name="oldName"/> to <paramref name="newName"/> on a rename.
    /// A chartsheet/dialogsheet has no such surface — its implementation is a
    /// no-op (references to it from other sheets are handled by the workbook).
    /// </summary>
    void RewriteSheetReferences(string oldName, string newName);

    /// <summary>
    /// Rewrites this sheet's references to a now-removed sheet into
    /// <c>#REF!</c> on a delete. No-op for a chartsheet/dialogsheet.
    /// </summary>
    void RewriteSheetReferencesToRefError(string removedName);
}
