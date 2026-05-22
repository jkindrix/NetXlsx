// XssfSheet — AutoFilter surface per decision I-56.
// Core class structure is in XssfSheet.cs.

using System;
using NPOI.SS.Util;

namespace NetXlsx;

internal sealed partial class XssfSheet
{
    public void SetAutoFilter(string a1Range)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Range);
        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        _underlying.SetAutoFilter(new CellRangeAddress(r1 - 1, r2 - 1, c1 - 1, c2 - 1));
    }

    public void ClearAutoFilter()
    {
        _workbook.ThrowIfDisposed();
        // NPOI 2.7.3 does not expose IsSet / unset accessors on
        // CT_Worksheet for the autoFilter element; setting the
        // property to null removes it from the serialized XML.
        // The auxiliary _FilterDatabase built-in name is left in
        // place — Excel tolerates a stale name pointing at an
        // absent autoFilter.
        _underlying.GetCTWorksheet().autoFilter = null;
    }

    public bool HasAutoFilter
    {
        get
        {
            _workbook.ThrowIfDisposed();
            return _underlying.GetCTWorksheet().autoFilter is not null;
        }
    }

    public string? AutoFilterRange
    {
        get
        {
            _workbook.ThrowIfDisposed();
            var af = _underlying.GetCTWorksheet().autoFilter;
            return af?.@ref;
        }
    }
}
