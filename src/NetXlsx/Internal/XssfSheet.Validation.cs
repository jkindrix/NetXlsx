// XssfSheet — data-validation surface per decision I-55.
// Core class structure is in XssfSheet.cs.

using System;
using NPOI.SS.Util;

namespace NetXlsx;

internal sealed partial class XssfSheet
{
    public void AddValidation(string a1Range, DataValidation validation)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Range);
        ArgumentNullException.ThrowIfNull(validation);

        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        var helper = _underlying.GetDataValidationHelper();
        var constraint = validation.Build(helper);
        var area = new CellRangeAddressList(r1 - 1, r2 - 1, c1 - 1, c2 - 1);
        var npoiValidation = helper.CreateValidation(constraint, area);
        _underlying.AddValidationData(npoiValidation);
    }
}
