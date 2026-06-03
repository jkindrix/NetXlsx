// I-82 engine swap — Open XML SDK-backed ITable (CF/validation/tables/
// autofilter/sort slice). Created via OoxmlSheet.AddTable; not
// public-constructible. Surface per decisions I-51 (tables) + I-64 (totals).
//
// A table is a TableDefinitionPart hung off the WorksheetPart; this wrapper
// holds the part and reaches its <table> root on demand. Totals-row
// SUBTOTAL/custom formulas are written into the sheet cells at the DOM level
// (<c><f>…</f></c>) — the public ICell.SetFormula is a later slice, but
// internal engine code may author the node directly, mirroring how the NPOI
// engine writes the cell via SetCellFormula.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlTable : ITable
{
    private readonly OoxmlWorkbook _workbook;
    private readonly OoxmlSheet _sheet;
    private readonly TableDefinitionPart _part;

    internal OoxmlTable(OoxmlWorkbook workbook, OoxmlSheet sheet, TableDefinitionPart part)
    {
        _workbook = workbook;
        _sheet = sheet;
        _part = part;
    }

    internal TableDefinitionPart Part => _part;

    // TableDefinitionPart.Table is annotated nullable; AddTable always sets it.
    private S.Table Table => _part.Table!;

    public string Name
    {
        get { _workbook.ThrowIfDisposed(); return Table.Name?.Value ?? string.Empty; }
    }

    public string DisplayName
    {
        get { _workbook.ThrowIfDisposed(); return Table.DisplayName?.Value ?? string.Empty; }
        set
        {
            _workbook.ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(value);
            Table.DisplayName = value;
        }
    }

    public string Address
    {
        get
        {
            _workbook.ThrowIfDisposed();
            var (r1, c1, r2, c2) = ParseReference();
            return CellAddress.FormatRange(r1, c1, r2, c2);
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
            // Refresh from the header row, mirroring NPOI's UpdateHeaders: a
            // non-empty string header cell renames its column; otherwise the
            // stored name stands.
            var (r1, c1, _, _) = ParseReference();
            var cols = Columns().ToList();
            for (int i = 0; i < cols.Count; i++)
            {
                var cell = _sheet.CellHandle(r1, c1 + i);
                if (cell.Kind == CellKind.String && cell.GetString() is { Length: > 0 } text)
                    cols[i].Name = text;
            }
            return cols.Select(c => c.Name?.Value ?? string.Empty).ToArray();
        }
    }

    public bool HasTotalsRow
    {
        get { _workbook.ThrowIfDisposed(); return (Table.TotalsRowCount?.Value ?? 0) > 0; }
    }

    public string? StyleName
    {
        get { _workbook.ThrowIfDisposed(); return Table.TableStyleInfo?.Name?.Value; }
        set
        {
            _workbook.ThrowIfDisposed();
            var info = Table.TableStyleInfo;
            if (value is null)
            {
                info?.Remove();
                return;
            }
            if (info is null)
            {
                // tableStyleInfo is the last CT_Table child before extLst.
                Table.TableStyleInfo = new S.TableStyleInfo { Name = value };
            }
            else
            {
                info.Name = value;
            }
        }
    }

    // Escape hatch divergence (I-82): no NPOI table exists on the SDK engine.
    public NPOI.XSSF.UserModel.XSSFTable Underlying => throw new NotSupportedException(
        "ITable.Underlying (NPOI XSSFTable) is not available on the Open XML SDK " +
        "engine (I-82). Use IWorkbook.OpenXmlDocument for the SDK escape hatch.");

    // ---- Totals row (decision I-64) -----------------------------------

    public void AddTotalsRow()
    {
        _workbook.ThrowIfDisposed();
        if (HasTotalsRow) return;

        var (r1, c1, r2, c2) = ParseReference();
        Table.Reference = CellAddress.FormatRange(r1, c1, r2 + 1, c2);
        Table.TotalsRowCount = 1u;
        // The NPOI engine trims a table <autoFilter> back to the data range
        // here; this engine never creates one (matching NPOI's CreateTable
        // output), so there is nothing to trim.
    }

    public void RemoveTotalsRow()
    {
        _workbook.ThrowIfDisposed();
        if (!HasTotalsRow) return;

        Table.TotalsRowCount = null;   // absent == 0 (the schema default)

        // Clear per-column functions, formulas, and labels — no stale
        // metadata once the totals row is gone.
        foreach (var col in Columns())
        {
            col.TotalsRowFunction = null;
            col.TotalsRowFormula = null;
            col.TotalsRowLabel = null;
        }

        // Blank the totals-row cells, then shrink the range by one row.
        var (r1, c1, r2, c2) = ParseReference();
        for (int col = c1; col <= c2; col++)
            _sheet.CellHandle(r2, col).Clear();
        Table.Reference = CellAddress.FormatRange(r1, c1, r2 - 1, c2);
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
        ctCol.TotalsRowFunction = ToTotalsFunction(function);
        ctCol.TotalsRowFormula = null;
        ctCol.TotalsRowLabel = null;

        if (function == TotalsRowFunction.None)
        {
            ClearTotalsCell(columnIndex);
            return;
        }
        // Write the SUBTOTAL formula directly into the cell so the total
        // renders in any conforming viewer, not just one that auto-populates
        // from totalsRowFunction metadata on open (NPOI-engine parity).
        int code = SubtotalCodeFor(function);
        var formula = $"SUBTOTAL({code.ToString(CultureInfo.InvariantCulture)},{Name}[{columnName}])";
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
        ctCol.TotalsRowFunction = S.TotalsRowFunctionValues.Custom;
        ctCol.TotalsRowFormula = new S.TotalsRowFormula(body);
        ctCol.TotalsRowLabel = null;

        WriteTotalsCell(columnIndex, body, isFormula: true);
    }

    public void SetColumnTotalLabel(string columnName, string label)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(columnName);
        ArgumentNullException.ThrowIfNull(label);
        EnsureTotalsRow();

        var (ctCol, columnIndex) = ResolveColumn(columnName);
        ctCol.TotalsRowLabel = label;
        ctCol.TotalsRowFunction = null;   // absent == none
        ctCol.TotalsRowFormula = null;

        WriteTotalsCell(columnIndex, label, isFormula: false);
    }

    // ---- helpers --------------------------------------------------------

    private (int r1, int c1, int r2, int c2) ParseReference()
        => CellAddress.ParseRange(Table.Reference?.Value!);

    private IEnumerable<S.TableColumn> Columns()
        => Table.TableColumns?.Elements<S.TableColumn>() ?? Enumerable.Empty<S.TableColumn>();

    private void EnsureTotalsRow()
    {
        if (!HasTotalsRow)
        {
            throw new InvalidOperationException(
                "Table has no totals row. Call AddTotalsRow() first.");
        }
    }

    private (S.TableColumn ctCol, int columnIndex) ResolveColumn(string columnName)
    {
        int i = 0;
        foreach (var col in Columns())
        {
            if (string.Equals(col.Name?.Value, columnName, StringComparison.OrdinalIgnoreCase))
                return (col, i);
            i++;
        }
        throw new ArgumentException(
            $"Column '{columnName}' is not part of table '{Name}'.",
            nameof(columnName));
    }

    private void WriteTotalsCell(int columnIndexInTable, string content, bool isFormula)
    {
        var (_, c1, r2, _) = ParseReference();   // totals row is the current end row
        var cell = _sheet.GetOrCreateCell(r2, c1 + columnIndexInTable);
        if (isFormula)
        {
            cell.RemoveAllChildren();
            cell.DataType = null;
            cell.CellFormula = new S.CellFormula(content);
        }
        else
        {
            _sheet.CellHandle(r2, c1 + columnIndexInTable).SetString(content);
        }
    }

    private void ClearTotalsCell(int columnIndexInTable)
    {
        var (_, c1, r2, _) = ParseReference();
        _sheet.CellHandle(r2, c1 + columnIndexInTable).Clear();
    }

    private static S.TotalsRowFunctionValues ToTotalsFunction(TotalsRowFunction f) => f switch
    {
        TotalsRowFunction.Sum => S.TotalsRowFunctionValues.Sum,
        TotalsRowFunction.Min => S.TotalsRowFunctionValues.Minimum,
        TotalsRowFunction.Max => S.TotalsRowFunctionValues.Maximum,
        TotalsRowFunction.Average => S.TotalsRowFunctionValues.Average,
        TotalsRowFunction.Count => S.TotalsRowFunctionValues.Count,
        TotalsRowFunction.CountNumbers => S.TotalsRowFunctionValues.CountNumbers,
        TotalsRowFunction.StdDev => S.TotalsRowFunctionValues.StandardDeviation,
        TotalsRowFunction.Var => S.TotalsRowFunctionValues.Variance,
        TotalsRowFunction.Custom => S.TotalsRowFunctionValues.Custom,
        _ => S.TotalsRowFunctionValues.None,
    };

    // 100-series ignores rows hidden by AutoFilter — the right default for
    // table totals (decision I-64; identical to the NPOI engine's table).
    private static int SubtotalCodeFor(TotalsRowFunction f) => f switch
    {
        TotalsRowFunction.Average => 101,
        TotalsRowFunction.CountNumbers => 102,
        TotalsRowFunction.Count => 103,
        TotalsRowFunction.Max => 104,
        TotalsRowFunction.Min => 105,
        TotalsRowFunction.StdDev => 107,
        TotalsRowFunction.Sum => 109,
        TotalsRowFunction.Var => 110,
        _ => 0,
    };
}
