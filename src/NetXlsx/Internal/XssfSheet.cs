using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfSheet : ISheet
{
    private readonly XssfWorkbook _workbook;
    private readonly XSSFSheet _underlying;

    public XssfSheet(XssfWorkbook workbook, XSSFSheet underlying)
    {
        _workbook = workbook;
        _underlying = underlying;
    }

    public string Name
    {
        get { _workbook.ThrowIfDisposed(); return _underlying.SheetName; }
    }

    public IWorkbook Workbook
    {
        get { _workbook.ThrowIfDisposed(); return _workbook; }
    }

    public ICell this[string a1]
    {
        get
        {
            _workbook.ThrowIfDisposed();
            var (row1, col1) = CellAddress.Parse(a1);
            return GetOrCreateCell(row1, col1);
        }
    }

    public ICell this[int row, int column]
    {
        get
        {
            _workbook.ThrowIfDisposed();
            if (row < 1 || row > CellAddress.MaxRow)
                throw new ArgumentOutOfRangeException(nameof(row), row,
                    $"row must be in [1, {CellAddress.MaxRow}]");
            if (column < 1 || column > CellAddress.MaxColumn)
                throw new ArgumentOutOfRangeException(nameof(column), column,
                    $"column must be in [1, {CellAddress.MaxColumn}]");
            return GetOrCreateCell(row, column);
        }
    }

    public IRange Range(string a1Range)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Range);
        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        return new XssfRange(_workbook, this, r1, c1, r2, c2);
    }

    public IRange Range(int row1, int col1, int row2, int col2)
    {
        _workbook.ThrowIfDisposed();
        ValidateGridCoordinate(row1, col1);
        ValidateGridCoordinate(row2, col2);
        // Normalize so callers can pass corners in any order.
        if (row1 > row2) (row1, row2) = (row2, row1);
        if (col1 > col2) (col1, col2) = (col2, col1);
        return new XssfRange(_workbook, this, row1, col1, row2, col2);
    }

    private static void ValidateGridCoordinate(int row, int column)
    {
        if (row < 1 || row > CellAddress.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(row), row,
                $"row must be in [1, {CellAddress.MaxRow}]");
        if (column < 1 || column > CellAddress.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column,
                $"column must be in [1, {CellAddress.MaxColumn}]");
    }

    public IRow AppendRow()
    {
        _workbook.ThrowIfDisposed();
        int next0 = _underlying.PhysicalNumberOfRows == 0 ? 0 : _underlying.LastRowNum + 1;
        if (next0 + 1 > CellAddress.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(AppendRow),
                $"appending would exceed Excel's row limit of {CellAddress.MaxRow}");
        var npoiRow = (XSSFRow)_underlying.CreateRow(next0);
        return new XssfRow(_workbook, this, npoiRow);
    }

    public IRow Row(int index)
    {
        _workbook.ThrowIfDisposed();
        if (index < 1 || index > CellAddress.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(index), index,
                $"row index must be in [1, {CellAddress.MaxRow}]");
        int row0 = index - 1;
        var npoiRow = (XSSFRow?)_underlying.GetRow(row0) ?? (XSSFRow)_underlying.CreateRow(row0);
        return new XssfRow(_workbook, this, npoiRow);
    }

    public void FreezeRows(int rows)
    {
        _workbook.ThrowIfDisposed();
        FreezePane(rows, 0);
    }

    public void FreezeColumns(int cols)
    {
        _workbook.ThrowIfDisposed();
        FreezePane(0, cols);
    }

    public void FreezePane(int rows, int cols)
    {
        _workbook.ThrowIfDisposed();
        if (rows < 0) throw new ArgumentOutOfRangeException(nameof(rows), rows, "must be >= 0");
        if (cols < 0) throw new ArgumentOutOfRangeException(nameof(cols), cols, "must be >= 0");
        // NPOI: CreateFreezePane(colSplit, rowSplit) — column comes first.
        _underlying.CreateFreezePane(cols, rows);
    }

    public void MergeCells(string a1Range)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Range);
        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);

        // 1x1 merge is a no-op per decision I-38.
        if (r1 == r2 && c1 == c2) return;

        var newRegion = new CellRangeAddress(r1 - 1, r2 - 1, c1 - 1, c2 - 1);

        // Overlap check per design §6.4 — fail loud rather than silently
        // produce an invalid OOXML merge graph.
        for (int i = 0; i < _underlying.NumMergedRegions; i++)
        {
            var existing = _underlying.GetMergedRegion(i);
            if (RangesOverlap(existing, newRegion))
            {
                throw new InvalidOperationException(
                    $"MergeCells('{a1Range}') overlaps existing merged region " +
                    $"{CellAddress.FormatRange(existing.FirstRow + 1, existing.FirstColumn + 1, existing.LastRow + 1, existing.LastColumn + 1)}.");
            }
        }

        _underlying.AddMergedRegion(newRegion);
    }

    public void UnmergeCells(string a1Range)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Range);
        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        int row1_0 = r1 - 1, row2_0 = r2 - 1, col1_0 = c1 - 1, col2_0 = c2 - 1;

        for (int i = _underlying.NumMergedRegions - 1; i >= 0; i--)
        {
            var region = _underlying.GetMergedRegion(i);
            if (region.FirstRow == row1_0 && region.LastRow == row2_0
                && region.FirstColumn == col1_0 && region.LastColumn == col2_0)
            {
                _underlying.RemoveMergedRegion(i);
                return;
            }
        }
        // No matching region — silent no-op per design §6.4.
    }

    public IReadOnlyList<string> MergedRanges
    {
        get
        {
            _workbook.ThrowIfDisposed();
            int n = _underlying.NumMergedRegions;
            if (n == 0) return Array.Empty<string>();
            var list = new string[n];
            for (int i = 0; i < n; i++)
            {
                var r = _underlying.GetMergedRegion(i);
                list[i] = CellAddress.FormatRange(r.FirstRow + 1, r.FirstColumn + 1, r.LastRow + 1, r.LastColumn + 1);
            }
            return list;
        }
    }

    public bool Hidden
    {
        get
        {
            _workbook.ThrowIfDisposed();
            var idx = _workbook.Underlying.GetSheetIndex(_underlying);
            return _workbook.Underlying.GetSheetVisibility(idx) != SheetVisibility.Visible;
        }
        set
        {
            _workbook.ThrowIfDisposed();
            var idx = _workbook.Underlying.GetSheetIndex(_underlying);
            _workbook.Underlying.SetSheetVisibility(
                idx, value ? SheetVisibility.Hidden : SheetVisibility.Visible);
        }
    }

    public bool ShowGridlines
    {
        get { _workbook.ThrowIfDisposed(); return _underlying.DisplayGridlines; }
        set { _workbook.ThrowIfDisposed(); _underlying.DisplayGridlines = value; }
    }

    private static bool RangesOverlap(CellRangeAddress a, CellRangeAddress b) =>
        a.FirstRow <= b.LastRow && b.FirstRow <= a.LastRow
        && a.FirstColumn <= b.LastColumn && b.FirstColumn <= a.LastColumn;

    public XSSFSheet Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _underlying; }
    }

    private XssfCell GetOrCreateCell(int row1, int col1)
    {
        // NPOI rows/cols are 0-based; ours are 1-based per decision #3.
        int row0 = row1 - 1;
        int col0 = col1 - 1;

        // Materialize on access (decision #40 — every cell exists).
        var npoiRow = (XSSFRow?)_underlying.GetRow(row0) ?? (XSSFRow)_underlying.CreateRow(row0);
        var npoiCell = (XSSFCell?)npoiRow.GetCell(col0) ?? (XSSFCell)npoiRow.CreateCell(col0);
        return new XssfCell(_workbook, npoiCell, row1, col1);
    }
}
