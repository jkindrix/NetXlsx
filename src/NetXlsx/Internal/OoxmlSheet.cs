// I-82 engine swap — Open XML SDK-backed ISheet.
//
// Foundation slice (parallel-engine, late-cutover strategy; see design I-82):
// this stub knows only its name and its owning workbook. Every cell/row/range/
// drawing/style member throws NotYet(...) until its slice lands. The escape
// hatch (Underlying -> XSSFSheet) throws NotSupportedException — the SDK engine
// has no NPOI sheet to expose; the SDK document is reachable via
// IWorkbook.OpenXmlDocument.
//
// Member implementation tracks the slice order in the continuation plan:
// cells & rows -> styles -> rich text -> merges/panes/grouping -> drawings ->
// CF/validation/tables/autofilter/sort -> charts -> streaming.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlSheet : ISheet
{
    private readonly OoxmlWorkbook _workbook;
    private readonly WorksheetPart _worksheetPart;
    private string _name;

    internal OoxmlSheet(OoxmlWorkbook workbook, string name, WorksheetPart worksheetPart)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _name = name ?? throw new ArgumentNullException(nameof(name));
        _worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));
    }

    public string Name
    {
        get { _workbook.ThrowIfDisposed(); return _name; }
    }

    public IWorkbook Workbook
    {
        get { _workbook.ThrowIfDisposed(); return _workbook; }
    }

    internal OoxmlWorkbook WorkbookInternal => _workbook;

    // Allows the workbook to keep wrapper names in sync if a rename API lands.
    internal void SetNameInternal(string name) => _name = name;

    // ---- Grid access over the worksheet DOM --------------------------------
    // Excel requires <row> elements in ascending @r order and <c> elements in
    // ascending column order within a row. The Get-or-create helpers maintain
    // that ordering on insert; the Find helpers never mutate the DOM (decision
    // #40: reading a never-written address must not add nodes).

    // WorksheetPart.Worksheet is annotated nullable; create/open always set it.
    private S.Worksheet Worksheet => _worksheetPart.Worksheet!;

    private S.SheetData Data =>
        Worksheet.GetFirstChild<S.SheetData>() ?? Worksheet.AppendChild(new S.SheetData());

    internal S.Row? FindRow(int rowIndex)
    {
        foreach (var r in Data.Elements<S.Row>())
            if (r.RowIndex?.Value == (uint)rowIndex) return r;
        return null;
    }

    internal S.Cell? FindCell(int rowIndex, int col)
    {
        var row = FindRow(rowIndex);
        if (row is null) return null;
        foreach (var c in row.Elements<S.Cell>())
            if (ColumnOf(c) == col) return c;
        return null;
    }

    internal S.Row GetOrCreateRow(int rowIndex)
    {
        var data = Data;
        S.Row? successor = null;
        foreach (var r in data.Elements<S.Row>())
        {
            var ri = r.RowIndex?.Value ?? 0;
            if (ri == (uint)rowIndex) return r;
            if (ri > (uint)rowIndex) { successor = r; break; }
        }
        var newRow = new S.Row { RowIndex = (uint)rowIndex };
        if (successor is null) data.AppendChild(newRow);
        else data.InsertBefore(newRow, successor);
        return newRow;
    }

    internal S.Cell GetOrCreateCell(int rowIndex, int col)
    {
        var row = GetOrCreateRow(rowIndex);
        S.Cell? successor = null;
        foreach (var c in row.Elements<S.Cell>())
        {
            int cc = ColumnOf(c);
            if (cc == col) return c;
            if (cc > col) { successor = c; break; }
        }
        var newCell = new S.Cell { CellReference = CellAddress.Format(rowIndex, col) };
        if (successor is null) row.AppendChild(newCell);
        else row.InsertBefore(newCell, successor);
        return newCell;
    }

    internal OoxmlCell CellHandle(int row, int col) => new(this, row, col);

    // ---- Column-formatting access over <cols> ------------------------------
    // <cols> is column layout (width / hidden / default style). In schema order it
    // sits between <sheetFormatPr> and <sheetData>, so it is inserted before
    // <sheetData>. A <col> entry may span a [min,max] range; GetOrCreateColumn
    // splits a spanning entry so a single column gets its own attributes.

    private S.Columns GetOrCreateColumns()
        => OoxmlSchemaOrder.GetOrInsert(Worksheet, static () => new S.Columns());

    // Non-mutating lookup of the <col> entry covering a 1-based column, or null.
    internal S.Column? FindColumn(int col)
    {
        var cols = Worksheet.GetFirstChild<S.Columns>();
        if (cols is null) return null;
        uint c = (uint)col;
        foreach (var x in cols.Elements<S.Column>())
            if (x.Min?.Value <= c && c <= x.Max?.Value) return x;
        return null;
    }

    internal S.Column GetOrCreateColumn(int col)
    {
        var cols = GetOrCreateColumns();
        uint c = (uint)col;
        foreach (var existing in cols.Elements<S.Column>().ToList())
        {
            uint min = existing.Min?.Value ?? 0;
            uint max = existing.Max?.Value ?? 0;
            if (c < min || c > max) continue;
            if (min == max) return existing;
            return SplitColumn(cols, existing, c, min, max);
        }
        var fresh = new S.Column { Min = c, Max = c };
        InsertColumnOrdered(cols, fresh, c);
        return fresh;
    }

    // Splits a spanning <col> [min,max] so that `c` becomes its own single-column
    // entry, preserving the original's attributes on every produced fragment.
    private static S.Column SplitColumn(S.Columns cols, S.Column existing, uint c, uint min, uint max)
    {
        var mid = (S.Column)existing.CloneNode(true);
        mid.Min = c; mid.Max = c;
        if (c > min)
        {
            var left = (S.Column)existing.CloneNode(true);
            left.Min = min; left.Max = c - 1;
            cols.InsertBefore(left, existing);
        }
        cols.InsertBefore(mid, existing);
        if (c < max)
        {
            var right = (S.Column)existing.CloneNode(true);
            right.Min = c + 1; right.Max = max;
            cols.InsertBefore(right, existing);
        }
        existing.Remove();
        return mid;
    }

    private static void InsertColumnOrdered(S.Columns cols, S.Column fresh, uint c)
    {
        foreach (var x in cols.Elements<S.Column>())
        {
            if ((x.Min?.Value ?? 0) > c) { cols.InsertBefore(fresh, x); return; }
        }
        cols.AppendChild(fresh);
    }

    // Populated cells (value / inline string / formula present) within a
    // rectangle, in row-major then ascending-column order.
    internal IEnumerable<OoxmlCell> EnumeratePopulated(int r1, int c1, int r2, int c2)
    {
        foreach (var row in Data.Elements<S.Row>())
        {
            var ri = (int)(row.RowIndex?.Value ?? 0);
            if (ri < r1 || ri > r2) continue;
            foreach (var c in row.Elements<S.Cell>())
            {
                int col = ColumnOf(c);
                if (col < c1 || col > c2) continue;
                if (IsPopulated(c)) yield return new OoxmlCell(this, ri, col);
            }
        }
    }

    private static bool IsPopulated(S.Cell c) =>
        c.CellValue is not null || c.InlineString is not null || c.CellFormula is not null;

    private static int ColumnOf(S.Cell c)
    {
        var reference = c.CellReference?.Value;
        if (string.IsNullOrEmpty(reference)) return 0;
        var (_, col) = CellAddress.Parse(reference);
        return col;
    }

    private int MaxRowIndex()
    {
        int max = 0;
        foreach (var r in Data.Elements<S.Row>())
        {
            int ri = (int)(r.RowIndex?.Value ?? 0);
            if (ri > max) max = ri;
        }
        return max;
    }

    // ---- Not-yet-implemented surface (lands slice by slice; see I-82) -------

    private static NotImplementedException NotYet([CallerMemberName] string? member = null)
        => new(
            $"ISheet.{member} is not yet implemented on the Open XML SDK engine " +
            "(I-82 engine swap). It lands in a later slice; until then use the " +
            "legacy engine (Workbook.Create/Open) for this operation, or track " +
            "the swap in docs/design.md (I-82).");

    public ICell this[string a1]
    {
        get
        {
            _workbook.ThrowIfDisposed();
            var (row, col) = CellAddress.Parse(a1);
            return new OoxmlCell(this, row, col);
        }
    }

    public ICell this[int row, int column]
    {
        get
        {
            _workbook.ThrowIfDisposed();
            if (row < 1 || row > CellAddress.MaxRow)
                throw new ArgumentOutOfRangeException(nameof(row), row, $"row must be in [1, {CellAddress.MaxRow}]");
            if (column < 1 || column > CellAddress.MaxColumn)
                throw new ArgumentOutOfRangeException(nameof(column), column, $"column must be in [1, {CellAddress.MaxColumn}]");
            return new OoxmlCell(this, row, column);
        }
    }

    public IRange Range(string a1Range)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Range);
        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        return new OoxmlRange(this, r1, c1, r2, c2);
    }

    public IRange Range(int row1, int col1, int row2, int col2)
    {
        _workbook.ThrowIfDisposed();
        foreach (var (label, v, max) in new[] { ("row1", row1, CellAddress.MaxRow), ("col1", col1, CellAddress.MaxColumn), ("row2", row2, CellAddress.MaxRow), ("col2", col2, CellAddress.MaxColumn) })
            if (v < 1 || v > max)
                throw new ArgumentOutOfRangeException(label, v, $"{label} must be in [1, {max}]");
        return new OoxmlRange(this, row1, col1, row2, col2);
    }

    public IRow AppendRow()
    {
        _workbook.ThrowIfDisposed();
        int next = MaxRowIndex() + 1;
        GetOrCreateRow(next);
        return new OoxmlRow(this, next);
    }

    public IRow Row(int index)
    {
        _workbook.ThrowIfDisposed();
        if (index < 1 || index > CellAddress.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(index), index, $"row index must be in [1, {CellAddress.MaxRow}]");
        GetOrCreateRow(index);
        return new OoxmlRow(this, index);
    }

    public IColumn Column(int index)
    {
        _workbook.ThrowIfDisposed();
        if (index < 1 || index > CellAddress.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(index), index, $"column index must be in [1, {CellAddress.MaxColumn}]");
        return new OoxmlColumn(this, index);
    }

    public IColumn Column(string letter)
    {
        _workbook.ThrowIfDisposed();
        return new OoxmlColumn(this, CellAddress.ParseColumn(letter));
    }

    // Freeze / split panes, grouping (outline), sheet visibility, gridlines, and
    // default column width land in OoxmlSheet.Structure.cs (I-82 structure slice).
    // Sheet protection lands in OoxmlSheet.Protection.cs.

    public IChart AddChart(ChartType type, string startCell, string endCell, string categoryRange, string valueRange, string? title = null) => throw NotYet();

    // Pictures (AddPicture overloads + Pictures read-back) land in
    // OoxmlSheet.Pictures.cs; shapes/connectors (AddShape, AddConnector,
    // Connectors) land in OoxmlSheet.Shapes.cs (I-82 drawings slice).

    public void AddConditionalFormatting(string a1Range, params ConditionalFormat[] rules) => throw NotYet();
    public int ConditionalFormattingCount => throw NotYet();
    public void RemoveConditionalFormatting(int index) => throw NotYet();

    // SortRange lands in OoxmlSheet.Sort.cs (I-82 CF/validation/tables/
    // autofilter/sort slice).

    // Merges (MergeCells / MergeCellsStyled / UnmergeCells / MergedRanges) land
    // in OoxmlSheet.Merges.cs (I-82 structure slice).

    // Hidden / ShowGridlines / DefaultColumnWidth land in OoxmlSheet.Structure.cs.

    public ITable AddTable(string a1Range, string name, string? style = null) => throw NotYet();
    public IReadOnlyList<ITable> Tables => throw NotYet();
    public bool TryGetTable(string name, [MaybeNullWhen(false)] out ITable table) => throw NotYet();
    public void RemoveTable(ITable table) => throw NotYet();

    public void SetAutoFilter(string a1Range) => throw NotYet();
    public void ClearAutoFilter() => throw NotYet();
    public void SetAutoFilterColumn(int columnOffset, FilterCriteria criteria) => throw NotYet();
    public void ClearAutoFilterColumn(int columnOffset) => throw NotYet();
    public bool HasAutoFilter => throw NotYet();
    public string? AutoFilterRange => throw NotYet();

    public void AddValidation(string a1Range, DataValidation validation) => throw NotYet();

    // AddPicture overloads land in OoxmlSheet.Pictures.cs (I-82 drawings slice).

    // Protect / Unprotect / IsProtected land in OoxmlSheet.Protection.cs.

    // Escape hatch divergence (I-82): no NPOI sheet exists on the SDK engine.
    public NPOI.XSSF.UserModel.XSSFSheet Underlying => throw new NotSupportedException(
        "ISheet.Underlying (NPOI XSSFSheet) is not available on the Open XML SDK " +
        "engine (I-82). Use IWorkbook.OpenXmlDocument for the SDK escape hatch.");
}
