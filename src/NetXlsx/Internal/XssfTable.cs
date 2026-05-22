// XssfTable — internal wrapper around NPOI's XSSFTable.
// Created via XssfSheet.AddTable; not public-constructible.

using System;
using System.Collections.Generic;
using System.Globalization;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS;
using NPOI.SS.Util;
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

    // ---- Totals row (decision I-64) -----------------------------------

    public void AddTotalsRow()
    {
        _workbook.ThrowIfDisposed();
        if (HasTotalsRow) return;

        var area = _underlying.GetCellReferences();
        var start = area.FirstCell;
        var end = area.LastCell;

        // Extend the data range by one row downward.
        var extended = new AreaReference(
            start,
            new CellReference(end.Row + 1, end.Col),
            SpreadsheetVersion.EXCEL2007);
        _underlying.SetCellReferences(extended);

        _underlying.GetCTTable().totalsRowCount = 1;

        // AutoFilter on a table excludes the totals row from its
        // filterable range; NPOI's SetCellReferences updates
        // CT_AutoFilter.@ref to match the full range, so we trim it
        // back here.
        var ct = _underlying.GetCTTable();
        if (ct.autoFilter is not null)
        {
            var autoStart = new CellReference(start.Row, start.Col);
            var autoEnd = new CellReference(end.Row, end.Col);   // unchanged: pre-totals end row
            ct.autoFilter.@ref = new AreaReference(autoStart, autoEnd, SpreadsheetVersion.EXCEL2007).FormatAsString();
        }
    }

    public void RemoveTotalsRow()
    {
        _workbook.ThrowIfDisposed();
        if (!HasTotalsRow) return;

        var ct = _underlying.GetCTTable();
        ct.totalsRowCount = 0;

        // Clear per-column functions and labels — we don't carry stale
        // metadata once the totals row is gone.
        if (ct.tableColumns?.tableColumn is { } cols)
        {
            foreach (var c in cols)
            {
                c.totalsRowFunction = ST_TotalsRowFunction.none;
                c.totalsRowFormula = null;
                c.totalsRowLabel = null;
            }
        }

        // Clear the actual totals-row cell contents on the sheet. The
        // row stays addressable; we just blank the cells in the table's
        // column range so nothing stale shows up.
        var area = _underlying.GetCellReferences();
        var start = area.FirstCell;
        var end = area.LastCell;
        int totalsRow0 = end.Row;     // pre-shrink: totals row is current end
        var sheet = _underlying.GetXSSFSheet();
        var row = sheet.GetRow(totalsRow0);
        if (row is not null)
        {
            for (int col = start.Col; col <= end.Col; col++)
            {
                var cell = row.GetCell(col);
                cell?.SetBlank();
            }
        }

        // Shrink the data range by one row.
        var shrunk = new AreaReference(
            start,
            new CellReference(end.Row - 1, end.Col),
            SpreadsheetVersion.EXCEL2007);
        _underlying.SetCellReferences(shrunk);
    }

    public void SetColumnTotal(string columnName, TotalsRowFunction function)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(columnName);
        if (function == TotalsRowFunction.Custom)
        {
            throw new ArgumentException(
                "Use SetColumnTotal(name, customFormula) for TotalsRowFunction.Custom.",
                nameof(function));
        }
        EnsureTotalsRow();

        var (ctCol, columnIndex) = ResolveColumn(columnName);
        ctCol.totalsRowFunction = ToCT(function);
        ctCol.totalsRowFormula = null;
        ctCol.totalsRowLabel = null;

        // Write the SUBTOTAL formula directly into the cell so the
        // total renders in any conforming viewer, not just one that
        // auto-populates from totalsRowFunction metadata on open.
        if (function == TotalsRowFunction.None)
        {
            ClearTotalsCell(columnIndex);
            return;
        }
        var (subtotalCode, _) = SubtotalCodeFor(function);
        var dataRangeRef = StructuredColumnReference(columnName);
        var formula = $"SUBTOTAL({subtotalCode.ToString(CultureInfo.InvariantCulture)},{dataRangeRef})";
        WriteTotalsCell(columnIndex, formula, isFormula: true);
    }

    public void SetColumnTotal(string columnName, string customFormula)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(columnName);
        ArgumentNullException.ThrowIfNull(customFormula);
        EnsureTotalsRow();

        var body = customFormula.Length > 0 && customFormula[0] == '=' ? customFormula.Substring(1) : customFormula;
        if (body.Length == 0)
            throw new ArgumentException("Custom formula body is empty.", nameof(customFormula));

        var (ctCol, columnIndex) = ResolveColumn(columnName);
        ctCol.totalsRowFunction = ST_TotalsRowFunction.custom;
        ctCol.totalsRowFormula = new CT_TableFormula { Value = body };
        ctCol.totalsRowLabel = null;

        WriteTotalsCell(columnIndex, body, isFormula: true);
    }

    public void SetColumnTotalLabel(string columnName, string label)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(columnName);
        ArgumentNullException.ThrowIfNull(label);
        EnsureTotalsRow();

        var (ctCol, columnIndex) = ResolveColumn(columnName);
        ctCol.totalsRowLabel = label;
        ctCol.totalsRowFunction = ST_TotalsRowFunction.none;
        ctCol.totalsRowFormula = null;

        WriteTotalsCell(columnIndex, label, isFormula: false);
    }

    // ---- Totals helpers ----------------------------------------------

    private void EnsureTotalsRow()
    {
        if (!HasTotalsRow)
        {
            throw new InvalidOperationException(
                "Table has no totals row. Call AddTotalsRow() first.");
        }
    }

    private (CT_TableColumn ctCol, int columnIndex) ResolveColumn(string columnName)
    {
        _underlying.UpdateHeaders();
        var cols = _underlying.GetCTTable().tableColumns?.tableColumn;
        if (cols is not null)
        {
            for (int i = 0; i < cols.Count; i++)
            {
                if (string.Equals(cols[i].name, columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return (cols[i], i);
                }
            }
        }
        throw new ArgumentException(
            $"Column '{columnName}' is not part of table '{Name}'.",
            nameof(columnName));
    }

    private void WriteTotalsCell(int columnIndexInTable, string content, bool isFormula)
    {
        var area = _underlying.GetCellReferences();
        var start = area.FirstCell;
        var end = area.LastCell;
        int totalsRow0 = end.Row;
        int absoluteCol0 = start.Col + columnIndexInTable;

        var sheet = _underlying.GetXSSFSheet();
        var row = sheet.GetRow(totalsRow0) ?? sheet.CreateRow(totalsRow0);
        var cell = row.GetCell(absoluteCol0) ?? row.CreateCell(absoluteCol0);
        if (isFormula)
        {
            cell.SetCellFormula(content);
        }
        else
        {
            cell.SetCellValue(content);
        }
    }

    private void ClearTotalsCell(int columnIndexInTable)
    {
        var area = _underlying.GetCellReferences();
        var start = area.FirstCell;
        var end = area.LastCell;
        int totalsRow0 = end.Row;
        int absoluteCol0 = start.Col + columnIndexInTable;
        var row = _underlying.GetXSSFSheet().GetRow(totalsRow0);
        row?.GetCell(absoluteCol0)?.SetBlank();
    }

    /// <summary>
    /// Builds a structured-reference body for the data range of a
    /// table column — e.g. <c>Sales[Revenue]</c>. The structured form
    /// auto-tracks table growth; using it in the SUBTOTAL formula
    /// keeps totals correct after AddRow().
    /// </summary>
    private string StructuredColumnReference(string columnName)
    {
        // Excel quotes column names containing whitespace or
        // structured-reference special chars (#, [, ], ', @). v1.2
        // covers the common case (no special chars) without quoting;
        // names that would need quoting reach through Underlying.
        return $"{_underlying.Name}[{columnName}]";
    }

    private static ST_TotalsRowFunction ToCT(TotalsRowFunction f) => f switch
    {
        TotalsRowFunction.None => ST_TotalsRowFunction.none,
        TotalsRowFunction.Sum => ST_TotalsRowFunction.sum,
        TotalsRowFunction.Min => ST_TotalsRowFunction.min,
        TotalsRowFunction.Max => ST_TotalsRowFunction.max,
        TotalsRowFunction.Average => ST_TotalsRowFunction.average,
        TotalsRowFunction.Count => ST_TotalsRowFunction.count,
        TotalsRowFunction.CountNumbers => ST_TotalsRowFunction.countNums,
        TotalsRowFunction.StdDev => ST_TotalsRowFunction.stdDev,
        TotalsRowFunction.Var => ST_TotalsRowFunction.var,
        TotalsRowFunction.Custom => ST_TotalsRowFunction.custom,
        _ => ST_TotalsRowFunction.none,
    };

    private static (int subtotalCode, string functionName) SubtotalCodeFor(TotalsRowFunction f) => f switch
    {
        // 100-series ignores rows hidden by AutoFilter — the right
        // default for table totals (decision I-64).
        TotalsRowFunction.Average => (101, "AVERAGE"),
        TotalsRowFunction.CountNumbers => (102, "COUNT"),
        TotalsRowFunction.Count => (103, "COUNTA"),
        TotalsRowFunction.Max => (104, "MAX"),
        TotalsRowFunction.Min => (105, "MIN"),
        TotalsRowFunction.StdDev => (107, "STDEV"),
        TotalsRowFunction.Sum => (109, "SUM"),
        TotalsRowFunction.Var => (110, "VAR"),
        _ => (0, "NONE"),
    };
}
