// XssfSheet — sheet-protection surface per decision I-53.
// Core class structure is in XssfSheet.cs.

namespace NetXlsx;

internal sealed partial class XssfSheet
{
    public void Protect(string? password = null, SheetProtection? options = null)
    {
        _workbook.ThrowIfDisposed();
        var opts = options ?? SheetProtection.Default;

        if (password is not null)
        {
            // ProtectSheet with a non-null password creates the
            // sheetProtection element and hashes the password.
            _underlying.ProtectSheet(password);
        }
        else
        {
            // NPOI 2.7.3's ProtectSheet(null) *removes* protection (it
            // takes null to mean "unprotect"). To protect without a
            // password, mirror what ProtectSheet(non-null) does minus
            // the password step: create the CT_SheetProtection element
            // with sheet=true. Captured in implementation-notes.md.
            var ct = _underlying.GetCTWorksheet();
            var sp = ct.sheetProtection ?? ct.AddNewSheetProtection();
            sp.sheet = true;
            sp.scenarios = true;
            sp.objects = true;
        }

        // Granular flags — applied after ProtectSheet so the sheet-
        // protection element exists. Each method toggles a single
        // <sheetProtection> attribute on the underlying CT.
        _underlying.LockFormatCells(opts.LockFormatCells);
        _underlying.LockFormatColumns(opts.LockFormatColumns);
        _underlying.LockFormatRows(opts.LockFormatRows);
        _underlying.LockInsertColumns(opts.LockInsertColumns);
        _underlying.LockInsertRows(opts.LockInsertRows);
        _underlying.LockInsertHyperlinks(opts.LockInsertHyperlinks);
        _underlying.LockDeleteColumns(opts.LockDeleteColumns);
        _underlying.LockDeleteRows(opts.LockDeleteRows);
        _underlying.LockSelectLockedCells(opts.LockSelectLockedCells);
        _underlying.LockSelectUnlockedCells(opts.LockSelectUnlockedCells);
        _underlying.LockSort(opts.LockSort);
        _underlying.LockAutoFilter(opts.LockAutoFilter);
        _underlying.LockPivotTables(opts.LockPivotTables);
        _underlying.LockObjects(opts.LockObjects);
        _underlying.LockScenarios(opts.LockScenarios);
    }

    public void Unprotect()
    {
        _workbook.ThrowIfDisposed();
        // ProtectSheet(null) on an already-unprotected sheet is a no-op
        // upstream, but on a protected sheet it clears the protection
        // element. NPOI 2.7.3 does not expose an explicit "unprotect"
        // method; passing null is the canonical pattern.
        _underlying.ProtectSheet(null);
    }

    public bool IsProtected
    {
        get { _workbook.ThrowIfDisposed(); return _underlying.IsSheetLocked; }
    }
}
