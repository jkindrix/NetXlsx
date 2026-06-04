// I-82 engine swap — Open XML SDK-backed ISheet.
//
// Grown slice-by-slice under the parallel-engine, late-cutover strategy
// (design I-82); since the v2.0.0 cutover this is THE engine. The escape
// hatch (Underlying) hands out the worksheet DOM root; the document is
// reachable via IWorkbook.Underlying.
//
// Member implementation tracks the slice order in the continuation plan:
// cells & rows -> styles -> rich text -> merges/panes/grouping -> drawings ->
// CF/validation/tables/autofilter/sort -> charts -> streaming.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
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
    //
    // Row lookups are served by a row-index -> <row> element cache (decision
    // I-87). Without it, GetOrCreateRow and MaxRowIndex linearly scan every
    // <row> on each call — ~12 full scans per appended 10-cell row — making
    // bulk DOM writes O(n²) (the v2.0.0 regression: Write5kRows 251 ms ->
    // 3,652 ms). Coherence model:
    //   * The engine itself never removes or renumbers a <row>; every engine
    //     row insert goes through GetOrCreateRow, which maintains the cache.
    //   * Every public escape-hatch getter (IWorkbook/ISheet/IRow/ICell
    //     .Underlying) invalidates ALL sheets' caches (any DOM node reaches
    //     the whole package via part traversal), so the acquire-then-mutate
    //     pattern is always coherent — the ThemeInfo cache precedent.
    //   * Backstop: every cache hit is liveness-checked (still parented under
    //     the current <sheetData>, @r unchanged) and the append/max fast paths
    //     verify the cached max against the live tail — a stale entry triggers
    //     a one-shot rebuild instead of returning a detached node.
    //   * Out-of-contract structural mutation through a STORED reference,
    //     interleaved with facade calls, is documented on the hatches:
    //     re-acquire any Underlying member after such mutations.

    // WorksheetPart.Worksheet is annotated nullable; create/open always set it.
    private S.Worksheet Worksheet => _worksheetPart.Worksheet!;

    private S.SheetData Data =>
        Worksheet.GetFirstChild<S.SheetData>() ?? Worksheet.AppendChild(new S.SheetData());

    // Row-index -> element cache (I-87). Null = invalidated; rebuilt lazily by
    // one scan. _cachedMaxRow mirrors the largest explicit @r in the cache.
    private Dictionary<int, S.Row>? _rowCache;
    private int _cachedMaxRow;

    internal void InvalidateRowCache() => _rowCache = null;

    private Dictionary<int, S.Row> EnsureRowCache(S.SheetData data)
    {
        if (_rowCache is not null) return _rowCache;
        var map = new Dictionary<int, S.Row>();
        int max = 0;
        foreach (var r in data.Elements<S.Row>())
        {
            // Rows without an explicit @r were invisible to the pre-cache
            // FindRow/MaxRowIndex scans (null never equals a 1-based index);
            // keep them invisible. Duplicate @r (malformed input): first in
            // document order wins, matching the pre-cache scan.
            if (r.RowIndex?.Value is not uint ri || ri == 0) continue;
            map.TryAdd((int)ri, r);
            if ((int)ri > max) max = (int)ri;
        }
        _cachedMaxRow = max;
        return _rowCache = map;
    }

    // A cached entry is trusted only while it is still parented under the
    // current <sheetData> with an unchanged @r — the backstop against
    // out-of-contract mutation through a stored escape-hatch reference.
    private static bool IsLive(S.Row row, S.SheetData data, int rowIndex) =>
        ReferenceEquals(row.Parent, data) && row.RowIndex?.Value == (uint)rowIndex;

    // O(1) guard for the append fast path: the cached max is trusted only if
    // the live tail does not contradict it (an out-of-contract append would
    // leave a <row> with @r >= rowIndex at the tail). A tail without explicit
    // @r matches pre-cache semantics (such rows were never successors).
    private static bool AppendTailTrusted(S.SheetData data, int rowIndex)
    {
        var last = data.LastChild;
        if (last is null) return true;
        if (last is not S.Row tail) return false;
        return tail.RowIndex?.Value is not uint tri || tri < (uint)rowIndex;
    }

    internal S.Row? FindRow(int rowIndex)
    {
        var data = Data;
        var cache = EnsureRowCache(data);
        if (cache.TryGetValue(rowIndex, out var row))
        {
            if (IsLive(row, data, rowIndex)) return row;
            InvalidateRowCache();
            cache = EnsureRowCache(data);
            return cache.TryGetValue(rowIndex, out row) ? row : null;
        }
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
        var cache = EnsureRowCache(data);
        if (cache.TryGetValue(rowIndex, out var cached))
        {
            if (IsLive(cached, data, rowIndex)) return cached;
            InvalidateRowCache();
            cache = EnsureRowCache(data);
            if (cache.TryGetValue(rowIndex, out cached)) return cached;
        }

        if (rowIndex > _cachedMaxRow && AppendTailTrusted(data, rowIndex))
        {
            var appended = new S.Row { RowIndex = (uint)rowIndex };
            data.AppendChild(appended);
            cache[rowIndex] = appended;
            _cachedMaxRow = rowIndex;
            return appended;
        }

        // General path (mid-grid insert, or an untrusted tail): one ordered
        // scan — adopt an out-of-band row if the index exists in the DOM,
        // else insert before the first higher-indexed sibling.
        S.Row? successor = null;
        foreach (var r in data.Elements<S.Row>())
        {
            var ri = r.RowIndex?.Value ?? 0;
            if (ri == (uint)rowIndex) { cache[rowIndex] = r; return r; }
            if (ri > (uint)rowIndex) { successor = r; break; }
        }
        var newRow = new S.Row { RowIndex = (uint)rowIndex };
        if (successor is null) data.AppendChild(newRow);
        else data.InsertBefore(newRow, successor);
        cache[rowIndex] = newRow;
        if (rowIndex > _cachedMaxRow) _cachedMaxRow = rowIndex;
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
        var data = Data;
        EnsureRowCache(data);
        // O(1) tail verification (the AppendTailTrusted counterpart): an
        // out-of-contract append leaves a tail whose explicit @r disagrees
        // with the cached max — rebuild rather than hand back a stale max.
        if (data.LastChild is S.Row tail && tail.RowIndex?.Value is uint tri
            && (int)tri != _cachedMaxRow)
        {
            InvalidateRowCache();
            EnsureRowCache(data);
        }
        return _cachedMaxRow;
    }

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
            ValidateGridCoordinate(row, column);
            return new OoxmlCell(this, row, column);
        }
    }

    // Effective caps: min(user-configured option, Excel hard cap). The option
    // defaults to the Excel hard cap, so default behavior is unchanged;
    // configuring a smaller value fails earlier (mirrors XssfSheet).
    private int EffectiveMaxRow =>
        Math.Min(_workbook.Options.MaxRowsPerSheet, CellAddress.MaxRow);
    private int EffectiveMaxColumn =>
        Math.Min(_workbook.Options.MaxColsPerSheet, CellAddress.MaxColumn);

    private void ValidateGridCoordinate(int row, int column)
    {
        int rowCap = EffectiveMaxRow;
        int colCap = EffectiveMaxColumn;
        if (row < 1 || row > rowCap)
            throw new ArgumentOutOfRangeException(nameof(row), row,
                $"row must be in [1, {rowCap}]");
        if (column < 1 || column > colCap)
            throw new ArgumentOutOfRangeException(nameof(column), column,
                $"column must be in [1, {colCap}]");
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
        ValidateGridCoordinate(row1, col1);
        ValidateGridCoordinate(row2, col2);
        return new OoxmlRange(this, row1, col1, row2, col2);
    }

    public IRow AppendRow()
    {
        _workbook.ThrowIfDisposed();
        int next = MaxRowIndex() + 1;
        int rowCap = EffectiveMaxRow;
        if (next > rowCap)
            throw new ArgumentOutOfRangeException(nameof(AppendRow),
                $"appending would exceed the configured row limit of {rowCap}");
        GetOrCreateRow(next);
        return new OoxmlRow(this, next);
    }

    public IRow Row(int index)
    {
        _workbook.ThrowIfDisposed();
        int rowCap = EffectiveMaxRow;
        if (index < 1 || index > rowCap)
            throw new ArgumentOutOfRangeException(nameof(index), index, $"row index must be in [1, {rowCap}]");
        GetOrCreateRow(index);
        return new OoxmlRow(this, index);
    }

    public int LastRowNumber
    {
        get
        {
            _workbook.ThrowIfDisposed();
            // Decision I-85: last row containing >=1 cell element. A row
            // materialized with no cells (Row(int)/AppendRow on an untouched
            // index) does not count; Clear() keeps the <c> node, so a cleared
            // cell still does.
            int last = 0;
            foreach (var r in Data.Elements<S.Row>())
            {
                if (!r.Elements<S.Cell>().Any()) continue;
                int ri = (int)(r.RowIndex?.Value ?? 0);
                if (ri > last) last = ri;
            }
            return last;
        }
    }

    public IColumn Column(int index)
    {
        _workbook.ThrowIfDisposed();
        int colCap = EffectiveMaxColumn;
        if (index < 1 || index > colCap)
            throw new ArgumentOutOfRangeException(nameof(index), index, $"column index must be in [1, {colCap}]");
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

    // AddChart lands in OoxmlSheet.Charts.cs (I-82 charts slice).

    // Pictures (AddPicture overloads + Pictures read-back) land in
    // OoxmlSheet.Pictures.cs; shapes/connectors (AddShape, AddConnector,
    // Connectors) land in OoxmlSheet.Shapes.cs (I-82 drawings slice).

    // Conditional formatting (AddConditionalFormatting /
    // ConditionalFormattingCount / RemoveConditionalFormatting) lands in
    // OoxmlSheet.ConditionalFormatting.cs (I-82 CF/validation/tables/
    // autofilter/sort slice).

    // SortRange lands in OoxmlSheet.Sort.cs (I-82 CF/validation/tables/
    // autofilter/sort slice).

    // Merges (MergeCells / MergeCellsStyled / UnmergeCells / MergedRanges) land
    // in OoxmlSheet.Merges.cs (I-82 structure slice).

    // Hidden / ShowGridlines / DefaultColumnWidth land in OoxmlSheet.Structure.cs.

    // Tables (AddTable / Tables / TryGetTable / RemoveTable) land in
    // OoxmlSheet.Tables.cs + OoxmlTable.cs (I-82 CF/validation/tables/
    // autofilter/sort slice).

    // AutoFilter (SetAutoFilter / ClearAutoFilter / per-column criteria /
    // HasAutoFilter / AutoFilterRange) lands in OoxmlSheet.AutoFilter.cs
    // (I-82 CF/validation/tables/autofilter/sort slice).

    // AddValidation lands in OoxmlSheet.Validation.cs (I-82 CF/validation/
    // tables/autofilter/sort slice).

    // AddPicture overloads land in OoxmlSheet.Pictures.cs (I-82 drawings slice).

    // Protect / Unprotect / IsProtected land in OoxmlSheet.Protection.cs.

    // Escape hatch (#32 / I-82): the worksheet DOM root. Disposal first.
    // Handing out the DOM invalidates the row caches (I-87) so mutations made
    // through the returned reference are observed by subsequent facade calls.
    public S.Worksheet Underlying
    {
        get
        {
            _workbook.ThrowIfDisposed();
            _workbook.InvalidateRowCaches();
            return Worksheet;
        }
    }
}
