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

internal sealed partial class OoxmlSheet : IOoxmlSheet
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
        get { ThrowIfUnusable(); return _name; }
    }

    public IWorkbook Workbook
    {
        get { ThrowIfUnusable(); return _workbook; }
    }

    /// <summary>Always <see cref="SheetKind.Worksheet"/> — this is the grid-backed sheet.</summary>
    public SheetKind Kind
    {
        get { ThrowIfUnusable(); return SheetKind.Worksheet; }
    }

    public OoxmlWorkbook WorkbookInternal => _workbook;

    // The workbook keeps the wrapper name in sync when RenameSheet (I-90)
    // commits a rename — never called outside that path.
    public void SetNameInternal(string name) => _name = name;

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

    // WorksheetPart.Worksheet is null when the part's XML is corrupt enough
    // that the SDK's DOM load produced no root (R-37 sibling — the deep-fuzz
    // sweep hit this through a bit-flipped worksheet part). Classify instead
    // of null-forgiving: every access path then surfaces the documented
    // MalformedFileException rather than an NRE.
    private S.Worksheet Worksheet =>
        _worksheetPart.Worksheet
        ?? throw new MalformedFileException(
            $"worksheet part for sheet '{_name}' has no worksheet root element (corrupt or non-spreadsheet content)");

    private S.SheetData Data =>
        Worksheet.GetFirstChild<S.SheetData>() ?? Worksheet.AppendChild(new S.SheetData());

    // Row-index -> element cache (I-87). Null = invalidated; rebuilt lazily by
    // one scan. _cachedMaxRow mirrors the largest explicit @r in the cache.
    private Dictionary<int, S.Row>? _rowCache;
    private int _cachedMaxRow;

    // Last-resolved-cell memo (I-87). Bulk write patterns touch the same
    // address consecutively (Set then Style; read-merge-write inside Style),
    // and within-row sibling walks pay a CellAddress.Parse per visited <c> —
    // the memo turns the repeat resolution into one reference comparison.
    // Trusted only while the node is parented under the live resolved row
    // (the same liveness discipline as the row cache); reset with it.
    private S.Cell? _lastCell;
    private int _lastCellRow, _lastCellCol;

    internal void InvalidateRowCache()
    {
        _rowCache = null;
        _lastCell = null;
    }

    private S.Cell MemoCell(int rowIndex, int col, S.Cell cell)
    {
        _lastCell = cell;
        _lastCellRow = rowIndex;
        _lastCellCol = col;
        return cell;
    }

    private Dictionary<int, S.Row> EnsureRowCache(S.SheetData data)
    {
        if (_rowCache is not null) return _rowCache;
        var map = new Dictionary<int, S.Row>();
        int max = 0;
        foreach (var r in data.Elements<S.Row>())
        {
            // Rows without an explicit @r: opened files have them inferred
            // and materialized at Open (R-14, NormalizeMissingReferences);
            // an @r-less row seen HERE was added through the escape hatch
            // post-open (out-of-contract) and stays invisible, matching the
            // pre-cache scans. Duplicate @r (malformed input): first in
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
        if (_lastCellRow == rowIndex && _lastCellCol == col
            && _lastCell is { } memo && ReferenceEquals(memo.Parent, row))
            return memo;
        foreach (var c in row.Elements<S.Cell>())
            if (ColumnOf(c) == col) return MemoCell(rowIndex, col, c);
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

        // Memo hit: consecutive operations on the same address skip the walk.
        if (_lastCellRow == rowIndex && _lastCellCol == col
            && _lastCell is { } memo && ReferenceEquals(memo.Parent, row))
            return memo;

        // Tail fast path (the within-row analogue of the row append path):
        // ascending-column writes append after the current tail with a single
        // ColumnOf instead of a sibling walk. A non-<c> tail (e.g. a row
        // <extLst> on an opened file) falls through to the general walk,
        // which appends after it exactly as the pre-I-87 code did.
        if (row.LastChild is S.Cell tail)
        {
            int tc = ColumnOf(tail);
            if (tc == col) return MemoCell(rowIndex, col, tail);
            if (tc < col)
            {
                var appended = new S.Cell { CellReference = CellAddress.Format(rowIndex, col) };
                row.AppendChild(appended);
                return MemoCell(rowIndex, col, appended);
            }
        }

        S.Cell? successor = null;
        foreach (var c in row.Elements<S.Cell>())
        {
            int cc = ColumnOf(c);
            if (cc == col) return MemoCell(rowIndex, col, c);
            if (cc > col) { successor = c; break; }
        }
        var newCell = new S.Cell { CellReference = CellAddress.Format(rowIndex, col) };
        if (successor is null) row.AppendChild(newCell);
        else row.InsertBefore(newCell, successor);
        return MemoCell(rowIndex, col, newCell);
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
            ThrowIfUnusable();
            var (row, col) = CellAddress.Parse(a1);
            return new OoxmlCell(this, row, col);
        }
    }

    public ICell this[int row, int column]
    {
        get
        {
            ThrowIfUnusable();
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
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(a1Range);
        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        return new OoxmlRange(this, r1, c1, r2, c2);
    }

    public IRange Range(int row1, int col1, int row2, int col2)
    {
        ThrowIfUnusable();
        ValidateGridCoordinate(row1, col1);
        ValidateGridCoordinate(row2, col2);
        return new OoxmlRange(this, row1, col1, row2, col2);
    }

    public IRow AppendRow()
    {
        ThrowIfUnusable();
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
        ThrowIfUnusable();
        int rowCap = EffectiveMaxRow;
        if (index < 1 || index > rowCap)
            throw new ArgumentOutOfRangeException(nameof(index), index, $"row index must be in [1, {rowCap}]");
        GetOrCreateRow(index);
        return new OoxmlRow(this, index);
    }

    // R-14: materialize row/@r and c/@r where a spec-legal third-party file
    // omitted them (ECMA-376: absent = previous + 1). Called once at Open,
    // before any read path runs. Inference is Excel-compatible — throwing
    // would reject files Excel opens fine. A reference that is PRESENT but
    // unparseable is left untouched (that is corruption, not omission —
    // I-83 fail-loud paths own it), and inference stops for the rest of
    // that row since the running column is no longer trustworthy.
    internal void NormalizeMissingReferences()
    {
        int prevRow = 0;
        foreach (var r in Data.Elements<S.Row>())
        {
            int ri;
            if (r.RowIndex?.Value is uint explicitRi && explicitRi != 0)
            {
                ri = (int)explicitRi;
            }
            else if (r.RowIndex is null && prevRow < CellAddress.MaxRow)
            {
                ri = prevRow + 1;
                r.RowIndex = (uint)ri;
            }
            else
            {
                // @r="0" (malformed) or inference past the row limit:
                // leave the row as it was (invisible, pre-R-14 behavior).
                prevRow = Math.Max(prevRow, (int)(r.RowIndex?.Value ?? 0));
                continue;
            }
            prevRow = ri;

            int prevCol = 0;
            foreach (var c in r.Elements<S.Cell>())
            {
                if (c.CellReference?.Value is string a1)
                {
                    if (!CellAddress.TryParse(a1, out _, out int col)) break; // corrupt @r — stop inferring here
                    prevCol = col;
                    continue;
                }
                if (prevCol >= CellAddress.MaxColumn) break;
                prevCol++;
                c.CellReference = CellAddress.Format(ri, prevCol);
            }
        }
        InvalidateRowCache();
    }

    // R-13: refresh this sheet's <dimension> from the live cell extent.
    // Called by the workbook's Save so consumers that trust the declared
    // extent (openpyxl's read-only/streaming mode refuses an "unsized"
    // worksheet) can size the sheet without scanning it; a stale dimension
    // on an opened file is corrected by the same pass. Empty sheet → "A1"
    // (Excel's own convention). Extent counts rows holding >=1 <c> (the
    // I-85 LastRowNumber rule) and columns from cell @r; @r-less rows and
    // cells stay invisible here exactly as they are to the reader (R-14).
    internal void UpdateDimension()
    {
        int minRow = 0, maxRow = 0, minCol = 0, maxCol = 0;
        foreach (var r in Data.Elements<S.Row>())
        {
            if (r.RowIndex?.Value is not uint ri || ri == 0) continue;
            bool anyCell = false;
            foreach (var c in r.Elements<S.Cell>())
            {
                if (c.CellReference?.Value is not string a1) continue;
                if (!CellAddress.TryParse(a1, out _, out int col)) continue;
                anyCell = true;
                if (minCol == 0 || col < minCol) minCol = col;
                if (col > maxCol) maxCol = col;
            }
            if (!anyCell) continue;
            if (minRow == 0 || (int)ri < minRow) minRow = (int)ri;
            if ((int)ri > maxRow) maxRow = (int)ri;
        }
        string reference = minRow == 0
            ? "A1"
            : minRow == maxRow && minCol == maxCol
                ? CellAddress.Format(minRow, minCol)
                : CellAddress.FormatRange(minRow, minCol, maxRow, maxCol);
        Worksheet.SheetDimension = new S.SheetDimension { Reference = reference };
    }

    public int LastRowNumber
    {
        get
        {
            ThrowIfUnusable();
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
        ThrowIfUnusable();
        int colCap = EffectiveMaxColumn;
        if (index < 1 || index > colCap)
            throw new ArgumentOutOfRangeException(nameof(index), index, $"column index must be in [1, {colCap}]");
        return new OoxmlColumn(this, index);
    }

    public IColumn Column(string letter)
    {
        ThrowIfUnusable();
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
            ThrowIfUnusable();
            _workbook.InvalidateRowCaches();
            return Worksheet;
        }
    }
}
