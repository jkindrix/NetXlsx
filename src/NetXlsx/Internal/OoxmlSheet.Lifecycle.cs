// I-90 sheet lifecycle (ledger R-12) — the sheet side: ISheet.Rename, the
// per-part-subtree reference rewrite the workbook's RenameSheet/RemoveSheet
// fan out to (rename → new name, delete → #REF!), and the removed-sheet
// access guard added by slice 2 (RemoveSheet).

using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xne = DocumentFormat.OpenXml.Office.Excel;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    // Workbook-side lifecycle code locates this sheet's <sheet> entry (via
    // SheetElementFor) and its part subtree through here.
    internal WorksheetPart WorksheetPartInternal => _worksheetPart;

    public void Rename(string newName)
    {
        ThrowIfUnusable();
        _workbook.RenameSheet(this, newName);
    }

    // ---- Removed-sheet access guard (I-90 slice 2) -------------------------
    // After IWorkbook.RemoveSheet detaches this sheet, every public ISheet
    // member (and, via _sheet, every OoxmlCell member) must throw — distinct
    // from the disposed-workbook ObjectDisposedException, matching the design's
    // removed-sheet-access language. The flag is one-way: a removed sheet never
    // returns to the workbook.

    private bool _removed;

    // Set by OoxmlWorkbook.RemoveSheet once the sheet's parts and entries have
    // been torn down. The wrapper instance lingers only as a tombstone.
    internal void MarkRemoved() => _removed = true;

    // The unified liveness guard the public surface routes through (replacing
    // the bare _workbook.ThrowIfDisposed()). Disposal is checked first so a
    // disposed workbook still surfaces ObjectDisposedException; a live workbook
    // with this sheet removed surfaces InvalidOperationException.
    internal void ThrowIfUnusable()
    {
        _workbook.ThrowIfDisposed();
        if (_removed)
            throw new InvalidOperationException(
                $"sheet '{_name}' has been removed from the workbook.");
    }

    /// <summary>
    /// Rename's per-sheet rewrite: replaces <paramref name="oldName"/> sheet
    /// references with the (normalized-quoted) <paramref name="newName"/>.
    /// </summary>
    internal void RewriteSheetReferences(string oldName, string newName)
        => RewriteSheetReferences(text => SheetReferenceLexer.Rewrite(text, oldName, newName));

    /// <summary>
    /// Delete's per-sheet rewrite (I-90 slice 2): rewrites every reference to
    /// the now-removed <paramref name="removedName"/> into Excel's <c>#REF!</c>
    /// dangling-reference literal across this sheet's formula-shaped surfaces —
    /// honest where leaving the dead name would silently corrupt. Internal
    /// hyperlink <c>@location</c> values pointing at the removed sheet are
    /// rewritten too, consistent with rename touching the same surface.
    /// </summary>
    internal void RewriteSheetReferencesToRefError(string removedName)
        => RewriteSheetReferences(text => SheetReferenceLexer.RewriteToRefError(text, removedName));

    /// <summary>
    /// Walks every formula-shaped surface owned by this sheet's part subtree
    /// and applies <paramref name="rewrite"/>: cell formulas (<c>&lt;f&gt;</c>,
    /// which covers shared-formula masters — followers carry no text), CF
    /// <c>&lt;formula&gt;</c>, DV <c>&lt;formula1/2&gt;</c>, <c>xm:f</c>
    /// (sparklines, plus any x14 CF/DV an opened file carries in
    /// <c>extLst</c>), internal hyperlink <c>@location</c>, table column
    /// formulas, and chart series references (<c>c:f</c>). Shared by the
    /// rename (new-name) and delete (<c>#REF!</c>) modes.
    /// </summary>
    private void RewriteSheetReferences(Func<string, string> rewrite)
    {
        var ws = Worksheet; // classified MalformedFileException on a corrupt part
        foreach (var f in ws.Descendants<S.CellFormula>()) RewriteLeafText(f, rewrite);
        foreach (var f in ws.Descendants<S.Formula>()) RewriteLeafText(f, rewrite);
        foreach (var f in ws.Descendants<S.Formula1>()) RewriteLeafText(f, rewrite);
        foreach (var f in ws.Descendants<S.Formula2>()) RewriteLeafText(f, rewrite);
        foreach (var f in ws.Descendants<Xne.Formula>()) RewriteLeafText(f, rewrite);

        // Internal hyperlink locations ("Sheet!A1", '#'-stripped at write).
        // Rewriting these deliberately EXCEEDS Excel, which leaves location
        // strings to break on rename (memo I-90 [A-2026-06-11]); on delete the
        // #REF! choice is the consistent one given rename already touches them.
        foreach (var h in ws.Descendants<S.Hyperlink>())
        {
            if (h.Location?.Value is not { Length: > 0 } location) continue;
            var rewritten = rewrite(location);
            if (!ReferenceEquals(rewritten, location)) h.Location = rewritten;
        }

        foreach (var tablePart in _worksheetPart.TableDefinitionParts)
        {
            if (tablePart.Table is not { } table) continue;
            foreach (var f in table.Descendants<S.CalculatedColumnFormula>()) RewriteLeafText(f, rewrite);
            foreach (var f in table.Descendants<S.TotalsRowFormula>()) RewriteLeafText(f, rewrite);
        }

        if (_worksheetPart.DrawingsPart is { } drawings)
        {
            foreach (var chartPart in drawings.ChartParts)
            {
                if (chartPart.ChartSpace is not { } chartSpace) continue;
                foreach (var f in chartSpace.Descendants<C.Formula>()) RewriteLeafText(f, rewrite);
            }
        }
    }

    private static void RewriteLeafText(OpenXmlLeafTextElement element, Func<string, string> rewrite)
    {
        var text = element.Text;
        if (text.Length == 0) return;
        var rewritten = rewrite(text);
        if (!ReferenceEquals(rewritten, text)) element.Text = rewritten;
    }
}
