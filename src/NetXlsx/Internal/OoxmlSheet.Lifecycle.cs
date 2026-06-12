// I-90 sheet lifecycle (slice 1: rename + move; ledger R-12) — the sheet
// side: ISheet.Rename plus the per-part-subtree reference rewrite the
// workbook's RenameSheet fans out to. RemoveSheet (slice 2) adds the
// removed-sheet access guard here.

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
        _workbook.ThrowIfDisposed();
        _workbook.RenameSheet(this, newName);
    }

    /// <summary>
    /// Rewrites <paramref name="oldName"/> sheet references to
    /// <paramref name="newName"/> in every formula-shaped surface owned by
    /// this sheet's part subtree: cell formulas (<c>&lt;f&gt;</c>, which
    /// covers shared-formula masters — followers carry no text), CF
    /// <c>&lt;formula&gt;</c>, DV <c>&lt;formula1/2&gt;</c>, <c>xm:f</c>
    /// (sparklines, plus any x14 CF/DV an opened file carries in
    /// <c>extLst</c>), internal hyperlink <c>@location</c>, table column
    /// formulas, and chart series references (<c>c:f</c>).
    /// </summary>
    internal void RewriteSheetReferences(string oldName, string newName)
    {
        var ws = Worksheet; // classified MalformedFileException on a corrupt part
        foreach (var f in ws.Descendants<S.CellFormula>()) RewriteLeafText(f, oldName, newName);
        foreach (var f in ws.Descendants<S.Formula>()) RewriteLeafText(f, oldName, newName);
        foreach (var f in ws.Descendants<S.Formula1>()) RewriteLeafText(f, oldName, newName);
        foreach (var f in ws.Descendants<S.Formula2>()) RewriteLeafText(f, oldName, newName);
        foreach (var f in ws.Descendants<Xne.Formula>()) RewriteLeafText(f, oldName, newName);

        // Internal hyperlink locations ("Sheet!A1", '#'-stripped at write).
        // Rewriting these deliberately EXCEEDS Excel, which leaves location
        // strings to break on rename (memo I-90 [A-2026-06-11]).
        foreach (var h in ws.Descendants<S.Hyperlink>())
        {
            if (h.Location?.Value is not { Length: > 0 } location) continue;
            var rewritten = SheetReferenceLexer.Rewrite(location, oldName, newName);
            if (!ReferenceEquals(rewritten, location)) h.Location = rewritten;
        }

        foreach (var tablePart in _worksheetPart.TableDefinitionParts)
        {
            if (tablePart.Table is not { } table) continue;
            foreach (var f in table.Descendants<S.CalculatedColumnFormula>()) RewriteLeafText(f, oldName, newName);
            foreach (var f in table.Descendants<S.TotalsRowFormula>()) RewriteLeafText(f, oldName, newName);
        }

        if (_worksheetPart.DrawingsPart is { } drawings)
        {
            foreach (var chartPart in drawings.ChartParts)
            {
                if (chartPart.ChartSpace is not { } chartSpace) continue;
                foreach (var f in chartSpace.Descendants<C.Formula>()) RewriteLeafText(f, oldName, newName);
            }
        }
    }

    private static void RewriteLeafText(OpenXmlLeafTextElement element, string oldName, string newName)
    {
        var text = element.Text;
        if (text.Length == 0) return;
        var rewritten = SheetReferenceLexer.Rewrite(text, oldName, newName);
        if (!ReferenceEquals(rewritten, text)) element.Text = rewritten;
    }
}
