// I-82 engine swap — Open XML SDK-backed INamedRange (structure slice).
//
// A defined name is a <definedName name="Sales" localSheetId="0">Sheet1!$A$1:
// $B$10</definedName> element in workbook.xml. This wrapper reads straight off
// that element: Name is @name, Formula is the element text (the refers-to body,
// no leading '='), and SheetScope resolves @localSheetId (a 0-based index into
// the workbook's sheet collection) back to the sheet name — null for a
// workbook-scoped name (no @localSheetId). Mirrors XssfNamedRange's contract.

using System;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlNamedRange : INamedRange
{
    private readonly OoxmlWorkbook _workbook;
    private readonly S.DefinedName _element;

    internal OoxmlNamedRange(OoxmlWorkbook workbook, S.DefinedName element)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _element = element ?? throw new ArgumentNullException(nameof(element));
    }

    public string Name
    {
        get { _workbook.ThrowIfDisposed(); return _element.Name?.Value ?? string.Empty; }
    }

    public string Formula
    {
        get { _workbook.ThrowIfDisposed(); return _element.Text ?? string.Empty; }
    }

    public string? SheetScope
    {
        get
        {
            _workbook.ThrowIfDisposed();
            if (_element.LocalSheetId?.Value is not uint local) return null;
            return _workbook.SheetNameAt((int)local);
        }
    }
}
