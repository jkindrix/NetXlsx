// XssfTable — internal wrapper around NPOI's XSSFTable.
// Created via XssfSheet.AddTable; not public-constructible.

using System.Collections.Generic;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfTable : ITable
{
    private readonly XssfWorkbook _workbook;
    private readonly XssfSheet _sheet;
    private readonly XSSFTable _underlying;

    public XssfTable(XssfWorkbook workbook, XssfSheet sheet, XSSFTable underlying)
    {
        _workbook = workbook;
        _sheet = sheet;
        _underlying = underlying;
    }

    public string Name
    {
        get { _workbook.ThrowIfDisposed(); return _underlying.Name; }
    }

    public string DisplayName
    {
        get { _workbook.ThrowIfDisposed(); return _underlying.DisplayName; }
        set
        {
            _workbook.ThrowIfDisposed();
            System.ArgumentNullException.ThrowIfNull(value);
            _underlying.DisplayName = value;
        }
    }

    public string Address
    {
        get
        {
            _workbook.ThrowIfDisposed();
            var area = _underlying.GetCellReferences();
            var start = area.FirstCell;
            var end = area.LastCell;
            return CellAddress.FormatRange(
                start.Row + 1, start.Col + 1,
                end.Row + 1, end.Col + 1);
        }
    }

    public ISheet Sheet
    {
        get { _workbook.ThrowIfDisposed(); return _sheet; }
    }

    public IReadOnlyList<string> ColumnNames
    {
        get
        {
            _workbook.ThrowIfDisposed();
            // UpdateHeaders re-reads cell values into CT_TableColumn.name,
            // keeping the snapshot in sync with header-cell edits.
            _underlying.UpdateHeaders();
            var cols = _underlying.GetColumns();
            if (cols == null || cols.Count == 0) return System.Array.Empty<string>();
            var list = new string[cols.Count];
            for (int i = 0; i < cols.Count; i++)
                list[i] = cols[i].Name ?? string.Empty;
            return list;
        }
    }

    public bool HasTotalsRow
    {
        get { _workbook.ThrowIfDisposed(); return _underlying.TotalsRowCount > 0; }
    }

    public string? StyleName
    {
        get { _workbook.ThrowIfDisposed(); return _underlying.StyleName; }
        set { _workbook.ThrowIfDisposed(); _underlying.StyleName = value; }
    }

    public XSSFTable Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _underlying; }
    }
}
