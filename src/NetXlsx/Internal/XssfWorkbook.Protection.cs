// XssfWorkbook — workbook-protection surface per decision I-54.
// Core class structure is in XssfWorkbook.cs.

namespace NetXlsx;

internal sealed partial class XssfWorkbook
{
    public void Protect(WorkbookProtection? options = null)
    {
        ThrowIfDisposed();
        var opts = options ?? WorkbookProtection.LockStructure;

        // NPOI's Lock*/Unlock* pair sets the flag on the
        // CT_WorkbookProtection element (creating it if absent).
        if (opts.Structure) _underlying.LockStructure(); else _underlying.UnlockStructure();
        if (opts.Windows)   _underlying.LockWindows();   else _underlying.UnlockWindows();
        if (opts.Revision)  _underlying.LockRevision();  else _underlying.UnlockRevision();
    }

    public void Unprotect()
    {
        ThrowIfDisposed();
        _underlying.UnlockStructure();
        _underlying.UnlockWindows();
        _underlying.UnlockRevision();
    }

    public bool IsProtected
    {
        get
        {
            ThrowIfDisposed();
            return _underlying.IsStructureLocked()
                || _underlying.IsWindowsLocked()
                || _underlying.IsRevisionLocked();
        }
    }
}
