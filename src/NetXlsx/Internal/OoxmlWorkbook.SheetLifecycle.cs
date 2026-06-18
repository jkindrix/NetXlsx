// I-90 sheet lifecycle (rename + move + delete; ledger R-12) — the S2 memo,
// signed off as amended 2026-06-11. Slice 1 landed Rename + MoveSheet; slice 2
// (this file's RemoveSheet) completes the issue.
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
    internal void RenameSheet(IOoxmlSheet sheet, string newName)
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
        SheetElementFor(sheet.SheetPartInternal).Name = newName;
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
        if (sheet is not IOoxmlSheet target || !ReferenceEquals(target.WorkbookInternal, this)
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
        var element = SheetElementFor(target.SheetPartInternal);
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

    /// <summary>
    /// Removes <paramref name="sheet"/> from the workbook (I-90 slice 2,
    /// completing R-12). Follows the <see cref="OoxmlSheet.RemoveTable"/>
    /// precedent (<see cref="ArgumentException"/> on a foreign or stale
    /// handle) plus the full delete contract: a workbook must keep at least one
    /// VISIBLE sheet; the worksheet part and its owned descendants, the
    /// <c>&lt;sheet&gt;</c> entry, the calcChain (wholesale), and pivot caches
    /// sourced from the sheet are torn down; defined names scoped to it are
    /// purged and later scopes re-indexed; cross-sheet references to it rewrite
    /// to <c>#REF!</c>; <c>activeTab</c> is clamped; and the wrapper (with its
    /// cells) becomes a tombstone whose members throw.
    /// </summary>
    public void RemoveSheet(ISheet sheet)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(sheet);
        using var _ = EnterMutation();
        if (sheet is not IOoxmlSheet target || !ReferenceEquals(target.WorkbookInternal, this)
            || !_sheetsByIndex.Contains(target))
        {
            throw new ArgumentException("sheet does not belong to this workbook.", nameof(sheet));
        }

        // A valid workbook must retain at least one VISIBLE sheet (count
        // visible, not total): a hidden sheet cannot be the last one standing.
        if (!_sheetsByIndex.Any(s => !ReferenceEquals(s, target) && !s.IsHiddenInternal))
        {
            throw new InvalidOperationException(
                "removing this sheet would leave the workbook with no visible sheet; " +
                "a workbook must contain at least one visible sheet.");
        }

        string removedName = target.Name;
        int removedPos = _sheetsByIndex.IndexOf(target);
        var before = _sheetsByIndex.ToArray();
        var part = target.SheetPartInternal;
        var wbPart = _document.WorkbookPart
            ?? throw new InvalidOperationException("Workbook has no workbook part.");

        // Drop the <sheet> entry (resolve it via the still-live relationship
        // first), then the worksheet part. The clone-based Save walks only the
        // reachable part graph, so once the part's relationship is gone its
        // whole subtree (drawings, comments, tables) is unreachable and dropped
        // from the output — and those descendants are reachable-at-open, never
        // captured as orphans, so ReinjectOrphanParts cannot resurrect them
        // (watch item (b); pinned by the zip-level no-orphan assert).
        var sheetElement = SheetElementFor(part);
        sheetElement.Remove();
        wbPart.DeletePart(part);

        // Drop the sheet from the lookups BEFORE the defined-name re-index, so a
        // name scoped to the now-removed position resolves to IndexOf == -1 and
        // is purged rather than re-pointed (watch item (c)).
        _sheetsByIndex.RemoveAt(removedPos);
        _sheetsByName.Remove(removedName);

        // calcChain is rebuilt by Excel and we never author one; an opened file
        // may carry it. Drop it WHOLESALE — calcChain's c/@i is a sheetId, NOT
        // a sheet position, so the wholesale delete makes a re-index neither
        // needed nor correct (memo [A-2026-06-11]; do not "fix" with a re-index
        // later).
        if (wbPart.CalculationChainPart is { } calcChain) wbPart.DeletePart(calcChain);

        RemovePivotCachesSourcedFrom(wbPart, removedName);
        PurgeDefinedNamesScopedTo(removedPos);
        ReindexLocalSheetIds(before);
        RewriteReferencesToRemovedSheet(removedName);

        // bookViews/@activeTab is CLAMPED into the shrunken range (the memo
        // asks only for clamping on delete — unlike MoveSheet, which makes the
        // tab FOLLOW the active sheet).
        ClampActiveTab();

        // The wrapper (and, via _sheet, its cells) becomes a tombstone: every
        // public member throws InvalidOperationException hereafter.
        target.MarkRemoved();
    }

    // Removes the workbook <pivotCache> entry and the cache definition part
    // (with its records-part child, via DeletePart) for every pivot cache
    // whose worksheetSource/@sheet named the removed sheet — the cache would
    // otherwise dangle. Cache definition parts hang off the workbook part
    // (the OOXML owner); a pivot TABLE that referenced a now-removed cache is
    // out of scope (it typically lived on the removed sheet and went with it).
    private static void RemovePivotCachesSourcedFrom(WorkbookPart wbPart, string removedName)
    {
        foreach (var cachePart in wbPart.GetPartsOfType<PivotTableCacheDefinitionPart>().ToList())
        {
            var source = cachePart.PivotCacheDefinition
                ?.GetFirstChild<S.CacheSource>()
                ?.GetFirstChild<S.WorksheetSource>();
            if (source?.Sheet?.Value is not { } sourceSheet
                || !string.Equals(sourceSheet, removedName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string rid = wbPart.GetIdOfPart(cachePart);
            if (wbPart.Workbook?.GetFirstChild<S.PivotCaches>() is { } caches)
            {
                foreach (var entry in caches.Elements<S.PivotCache>()
                             .Where(e => e.Id?.Value == rid).ToList())
                    entry.Remove();
                if (!caches.Elements<S.PivotCache>().Any()) caches.Remove();
            }
            wbPart.DeletePart(cachePart);
        }
    }

    // Removes every defined name scoped to the removed sheet's POSITION
    // (localSheetId == removedPos). Runs before ReindexLocalSheetIds so the
    // re-index never encounters a name pointing at the vanished sheet.
    private void PurgeDefinedNamesScopedTo(int removedPos)
    {
        var container = DefinedNamesContainer();
        if (container is null) return;
        foreach (var defined in container.Elements<S.DefinedName>().ToList())
        {
            if (defined.LocalSheetId?.Value == (uint)removedPos) defined.Remove();
        }
        if (!container.Elements<S.DefinedName>().Any()) container.Remove();
    }

    // Delete's counterpart to RewriteSheetReferences: rewrites every reference
    // to the removed sheet into #REF! across surviving sheets' part subtrees
    // and the workbook-owned defined-name bodies. Pivot caches are REMOVED
    // (above), not rewritten.
    private void RewriteReferencesToRemovedSheet(string removedName)
    {
        foreach (var sheet in _sheetsByIndex)
            sheet.RewriteSheetReferencesToRefError(removedName);

        if (DefinedNamesContainer() is { } names)
        {
            foreach (var defined in names.Elements<S.DefinedName>())
            {
                var body = defined.Text;
                if (body.Length == 0) continue;
                var rewritten = SheetReferenceLexer.RewriteToRefError(body, removedName);
                if (!ReferenceEquals(rewritten, body)) defined.Text = rewritten;
            }
        }
    }

    // Clamps bookViews/workbookView/@activeTab into the post-removal range.
    // Created workbooks carry no bookViews; only an existing out-of-range
    // attribute is rewritten.
    private void ClampActiveTab()
    {
        var views = _document.WorkbookPart?.Workbook?.GetFirstChild<S.BookViews>();
        if (views is null) return;
        int max = _sheetsByIndex.Count - 1;
        foreach (var view in views.Elements<S.WorkbookView>())
        {
            int active = (int)(view.ActiveTab?.Value ?? 0);
            int clamped = Math.Clamp(active, 0, Math.Max(max, 0));
            if (clamped != active) view.ActiveTab = (uint)clamped;
        }
    }

    // localSheetId is a zero-based sheet POSITION; re-point every
    // sheet-scoped defined name (including hidden built-ins like
    // _xlnm._FilterDatabase) at its sheet's new position.
    private void ReindexLocalSheetIds(IOoxmlSheet[] before)
    {
        var container = DefinedNamesContainer();
        if (container is null) return;
        foreach (var defined in container.Elements<S.DefinedName>())
        {
            if (defined.LocalSheetId?.Value is not uint lsid) continue;
            if (lsid >= (uint)before.Length) continue; // malformed out-of-range scope: leave as-is
            int now = _sheetsByIndex.IndexOf(before[lsid]);
            // now == -1 means the scoped sheet is gone (RemoveSheet purges such
            // names before this runs, so it shouldn't occur — but never write a
            // (uint)-1 scope if it slips through). Move never hits this branch.
            if (now < 0 || now == (int)lsid) continue;
            defined.LocalSheetId = (uint)now;
        }
    }

    // bookViews/workbookView/@activeTab is a zero-based position: keep the
    // sheet that was active active (clamping a malformed out-of-range value
    // into range first). Created workbooks carry no bookViews — nothing is
    // added; only an existing attribute whose effective value changed is
    // written.
    private void RetargetActiveTab(IOoxmlSheet[] before)
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
