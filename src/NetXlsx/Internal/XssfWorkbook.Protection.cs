// XssfWorkbook — workbook-protection surface per decisions I-54 + I-65.
// Core class structure is in XssfWorkbook.cs.

using System;
using NPOI.POIFS.Crypt;

namespace NetXlsx;

internal sealed partial class XssfWorkbook
{
    public void Protect(WorkbookProtection? options = null)
    {
        ProtectCore(options, password: null);
    }

    public void ProtectWithPassword(string password, WorkbookProtection? options = null)
    {
        ArgumentNullException.ThrowIfNull(password);
        ProtectCore(options, password);
    }

    private void ProtectCore(WorkbookProtection? options, string? password)
    {
        ThrowIfDisposed();
        var opts = options ?? WorkbookProtection.LockStructure;

        // NPOI's Lock*/Unlock* pair sets the flag on the
        // CT_WorkbookProtection element (creating it if absent).
        if (opts.Structure) _underlying.LockStructure(); else _underlying.UnlockStructure();
        if (opts.Windows)   _underlying.LockWindows();   else _underlying.UnlockWindows();
        if (opts.Revision)  _underlying.LockRevision();  else _underlying.UnlockRevision();

        // I-65: optional XOR-verifier password. NPOI 2.7.3 does not
        // expose a workbook-level SetPassword helper; write the
        // 16-bit verifier into CT_WorkbookProtection.workbookPassword
        // as a 2-byte big-endian array directly. The byte[] field
        // serializes via XmlHelper.WriteAttribute(byte[]) which emits
        // the OOXML-expected hex form.
        var ct = _underlying.GetCTWorkbook().workbookProtection;
        if (ct is not null)
        {
            if (password is null)
            {
                ct.workbookPassword = null;
            }
            else
            {
                int verifier = CryptoFunctions.CreateXorVerifier1(password);
                ct.workbookPassword = new byte[]
                {
                    (byte)((verifier >> 8) & 0xFF),
                    (byte)(verifier & 0xFF),
                };
            }
        }
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
