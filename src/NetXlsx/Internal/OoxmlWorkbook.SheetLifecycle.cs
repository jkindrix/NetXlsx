// I-90 sheet lifecycle (slice 1: rename + move; ledger R-12) — the S2 memo,
// signed off as amended 2026-06-11. RemoveSheet lands in slice 2.
//
// Rename is a METHOD (a Name setter was rejected): it validates per the
// AddSheet rules, can throw SheetNameException, and rewrites references
// document-wide via SheetReferenceLexer — side effects a property setter
// would hide. Rewritten surfaces: cell formulas on EVERY sheet, defined-name
// bodies INCLUDING the _xlnm.* built-ins (print areas/titles are just
// defined names — the memo amendment forbids filtering reserved names),
// internal hyperlink locations (exceeds Excel, deliberately), CF/DV
// formulas, chart c:f, pivot-cache worksheetSource/@sheet, sparkline xm:f,
// and table column formulas.
//
// MoveSheet takes the 1-based RESULTING position (remove-then-insert).
// Defined-name localSheetId is a zero-based sheet POSITION — not a sheetId —
// so every move re-indexes it; bookViews/@activeTab follows the sheet that
// was active and clamps a malformed out-of-range value.

using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlWorkbook
{
    /// <summary>
    /// Commits a sheet rename: validates the new name (the AddSheet rules,
    /// including R-9's tightened set and the I-88 control-character
    /// fail-fast), keeps <c>_sheetsByName</c> / the wrapper / the
    /// <c>&lt;sheet&gt;</c> entry coherent, and rewrites the old name
    /// document-wide. An exact-equal rename is a no-op; a case-only rename
    /// proceeds (references are rewritten to the new casing, as Excel does).
    /// </summary>
    internal void RenameSheet(OoxmlSheet sheet, string newName)
    {
        using var _ = EnterMutation();
        Workbook.ValidateSheetName(newName);
        string oldName = sheet.Name;
        if (string.Equals(oldName, newName, StringComparison.Ordinal)) return;
        if (_sheetsByName.TryGetValue(newName, out var existing) && !ReferenceEquals(existing, sheet))
            throw new SheetNameException(newName, "a sheet with this name already exists (case-insensitive)");

        // The rewrite below walks every sheet's DOM. It cannot trip the
        // corrupt-part MalformedFileException mid-walk: Open already forced
        // every worksheet root through the classified getter (R-14's
        // NormalizeMissingReferences), so a load-failing part never reaches
        // a live workbook — the rename is all-or-nothing.
        SheetElementFor(sheet.WorksheetPartInternal).Name = newName;
        _sheetsByName.Remove(oldName);
        _sheetsByName[newName] = sheet;
        sheet.SetNameInternal(newName);

        RewriteSheetReferences(oldName, newName);
    }

    public void MoveSheet(ISheet sheet, int newIndex)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(sheet);
        using var _ = EnterMutation();
        if (sheet is not OoxmlSheet target || !ReferenceEquals(target.WorkbookInternal, this)
            || !_sheetsByIndex.Contains(target))
        {
            throw new ArgumentException("sheet does not belong to this workbook.", nameof(sheet));
        }
        if (newIndex < 1 || newIndex > _sheetsByIndex.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(newIndex), newIndex,
                $"newIndex must be in [1, {_sheetsByIndex.Count}] (the 1-based resulting position).");
        }

        int oldPos = _sheetsByIndex.IndexOf(target);
        int newPos = newIndex - 1;
        if (newPos == oldPos) return;

        var before = _sheetsByIndex.ToArray();
        _sheetsByIndex.RemoveAt(oldPos);
        _sheetsByIndex.Insert(newPos, target);

        // Mirror the wrapper order in workbook.xml. The remaining <sheet>
        // siblings sit in the same relative order as the wrapper list did
        // after its RemoveAt, so "insert before the element at newPos"
        // reproduces List.Insert exactly.
        var sheetsElement = _document.WorkbookPart?.Workbook?.GetFirstChild<S.Sheets>()
            ?? throw new InvalidOperationException("Workbook has no <sheets> element.");
        var element = SheetElementFor(target.WorksheetPartInternal);
        element.Remove();
        var siblings = sheetsElement.Elements<S.Sheet>().ToList();
        if (newPos >= siblings.Count) sheetsElement.AppendChild(element);
        else sheetsElement.InsertBefore(element, siblings[newPos]);

        ReindexLocalSheetIds(before);
        RetargetActiveTab(before);

        // xl/calcChain.xml is deliberately NOT touched: calcChain's c/@i is
        // a sheetId — stable across reorder — not a sheet position, so a
        // move never invalidates it (memo I-90 [A-2026-06-11]; do not "fix"
        // this with a re-index later).
    }

    // localSheetId is a zero-based sheet POSITION; re-point every
    // sheet-scoped defined name (including hidden built-ins like
    // _xlnm._FilterDatabase) at its sheet's new position.
    private void ReindexLocalSheetIds(OoxmlSheet[] before)
    {
        var container = DefinedNamesContainer();
        if (container is null) return;
        foreach (var defined in container.Elements<S.DefinedName>())
        {
            if (defined.LocalSheetId?.Value is not uint lsid) continue;
            if (lsid >= (uint)before.Length) continue; // malformed out-of-range scope: leave as-is
            int now = _sheetsByIndex.IndexOf(before[lsid]);
            if (now != (int)lsid) defined.LocalSheetId = (uint)now;
        }
    }

    // bookViews/workbookView/@activeTab is a zero-based position: keep the
    // sheet that was active active (clamping a malformed out-of-range value
    // into range first). Created workbooks carry no bookViews — nothing is
    // added; only an existing attribute whose effective value changed is
    // written.
    private void RetargetActiveTab(OoxmlSheet[] before)
    {
        var views = _document.WorkbookPart?.Workbook?.GetFirstChild<S.BookViews>();
        if (views is null) return;
        foreach (var view in views.Elements<S.WorkbookView>())
        {
            int active = (int)(view.ActiveTab?.Value ?? 0);
            int clamped = Math.Min(active, before.Length - 1);
            int now = _sheetsByIndex.IndexOf(before[clamped]);
            if (now != active) view.ActiveTab = (uint)now;
        }
    }

    /// <summary>
    /// Rename's document-wide rewrite: fans out to every sheet's part
    /// subtree, then covers the workbook-owned surfaces (defined names,
    /// pivot caches).
    /// </summary>
    private void RewriteSheetReferences(string oldName, string newName)
    {
        foreach (var sheet in _sheetsByIndex)
            sheet.RewriteSheetReferences(oldName, newName);

        if (DefinedNamesContainer() is { } names)
        {
            foreach (var defined in names.Elements<S.DefinedName>())
            {
                var body = defined.Text;
                if (body.Length == 0) continue;
                var rewritten = SheetReferenceLexer.Rewrite(body, oldName, newName);
                if (!ReferenceEquals(rewritten, body)) defined.Text = rewritten;
            }
        }

        // Pivot-cache worksheetSource/@sheet is a LITERAL sheet name (no
        // quoting). NetXlsx does not author pivots, but opened files carry
        // them — a missed rename dangles the cache (memo [A-2026-06-11]).
        // GetAllParts reaches every cache part regardless of how it is
        // related (workbook pivotCaches entry or a pivot table's rel).
        foreach (var part in _document.GetAllParts())
        {
            if (part is not PivotTableCacheDefinitionPart cachePart) continue;
            var source = cachePart.PivotCacheDefinition
                ?.GetFirstChild<S.CacheSource>()
                ?.GetFirstChild<S.WorksheetSource>();
            if (source?.Sheet?.Value is { } sourceSheet
                && string.Equals(sourceSheet, oldName, StringComparison.OrdinalIgnoreCase))
            {
                source.Sheet = newName;
            }
        }
    }
}
