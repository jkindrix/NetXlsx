// I-82 engine swap — Open XML SDK-backed workbook protection (formulas/
// comments/hyperlinks slice rider) per decisions I-54 + I-65, plus
// IsMacroEnabled.
//
// Mirrors the NPOI engine's XssfWorkbook.Protection contract: Protect sets the
// structure/windows/revision flags explicitly from the options (so the result
// never depends on the schema's per-attribute defaults — NPOI's Lock*/Unlock*
// net behavior, and the explicit lockWindows="0"/lockRevision="0" the oracle
// dump shows); an optional password is stored as the legacy 16-bit XOR verifier
// in @workbookPassword (the sheet-protection slice's LegacyPasswordHash — same
// classic Excel hash, a UX guard, not security). <workbookProtection> is a
// strict-sequence CT_Workbook child (between <workbookPr> and <bookViews>)
// placed by OoxmlSchemaOrder.
//
// Unprotect removes the element entirely — consistent with this engine's sheet
// Unprotect, and a (publicly invisible) tidy-up over NPOI, which leaves an
// all-flags-false stub element behind. IsProtected reads the three flags, so
// both shapes report false identically.

using System;
using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlWorkbook
{
    public void Protect(WorkbookProtection? options = null)
        => ProtectCore(options, password: null);

    public void ProtectWithPassword(string password, WorkbookProtection? options = null)
    {
        ArgumentNullException.ThrowIfNull(password);
        ProtectCore(options, password);
    }

    private void ProtectCore(WorkbookProtection? options, string? password)
    {
        ThrowIfDisposed();
        var opts = options ?? WorkbookProtection.LockStructure;

        var wp = OoxmlSchemaOrder.GetOrInsert(
            _document.WorkbookPart!.Workbook!, static () => new S.WorkbookProtection());
        wp.LockStructure = opts.Structure;
        wp.LockWindows = opts.Windows;
        wp.LockRevision = opts.Revision;
        // I-65: a passwordless Protect clears any prior verifier (NPOI parity).
        wp.WorkbookPassword = password is null
            ? null
            : HexBinaryValue.FromString(OoxmlSheet.LegacyPasswordHash(password));
    }

    public void Unprotect()
    {
        ThrowIfDisposed();
        _document.WorkbookPart?.Workbook?.GetFirstChild<S.WorkbookProtection>()?.Remove();
    }

    public bool IsProtected
    {
        get
        {
            ThrowIfDisposed();
            var wp = _document.WorkbookPart?.Workbook?.GetFirstChild<S.WorkbookProtection>();
            return wp is not null
                && ((wp.LockStructure?.Value ?? false)
                    || (wp.LockWindows?.Value ?? false)
                    || (wp.LockRevision?.Value ?? false));
        }
    }

    // The macro-enabled signal is the document's content type (.xlsm's
    // sheet.macroEnabled.main+xml vs .xlsx's sheet.main+xml) — exactly what
    // NPOI's IsMacroEnabled() checks. The SDK exposes it as DocumentType.
    public bool IsMacroEnabled
    {
        get
        {
            ThrowIfDisposed();
            return _document.DocumentType == SpreadsheetDocumentType.MacroEnabledWorkbook;
        }
    }
}
