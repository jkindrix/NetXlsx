using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed partial class XssfSheet : ISheet
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
            ValidateGridCoordinate(row, column);
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

    // Effective caps: min(user-configured option, Excel hard cap).
    // The option defaults to the Excel hard cap, so default behavior
    // is unchanged. Configuring a smaller value fails earlier.
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

    public IRow AppendRow()
    {
        _workbook.ThrowIfDisposed();
        int next0 = _underlying.PhysicalNumberOfRows == 0 ? 0 : _underlying.LastRowNum + 1;
        int rowCap = EffectiveMaxRow;
        if (next0 + 1 > rowCap)
            throw new ArgumentOutOfRangeException(nameof(AppendRow),
                $"appending would exceed the configured row limit of {rowCap}");
        var npoiRow = (XSSFRow)_underlying.CreateRow(next0);
        return new XssfRow(_workbook, this, npoiRow);
    }

    public IRow Row(int index)
    {
        _workbook.ThrowIfDisposed();
        int rowCap = EffectiveMaxRow;
        if (index < 1 || index > rowCap)
            throw new ArgumentOutOfRangeException(nameof(index), index,
                $"row index must be in [1, {rowCap}]");
        int row0 = index - 1;
        var npoiRow = (XSSFRow?)_underlying.GetRow(row0) ?? (XSSFRow)_underlying.CreateRow(row0);
        return new XssfRow(_workbook, this, npoiRow);
    }

    public IColumn Column(int index)
    {
        _workbook.ThrowIfDisposed();
        int colCap = EffectiveMaxColumn;
        if (index < 1 || index > colCap)
            throw new ArgumentOutOfRangeException(nameof(index), index,
                $"column index must be in [1, {colCap}]");
        return new XssfColumn(_workbook, this, index);
    }

    public IColumn Column(string letter)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(letter);
        int col = CellAddress.ParseColumn(letter);
        return new XssfColumn(_workbook, this, col);
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

    public void GroupRows(int startRow, int endRow)
    {
        _workbook.ThrowIfDisposed();
        if (startRow < 1) throw new ArgumentOutOfRangeException(nameof(startRow), startRow, "must be >= 1");
        if (endRow < 1) throw new ArgumentOutOfRangeException(nameof(endRow), endRow, "must be >= 1");
        if (startRow > endRow) throw new ArgumentOutOfRangeException(nameof(startRow), startRow, "startRow must be <= endRow");
        _underlying.GroupRow(startRow - 1, endRow - 1);
    }

    public void UngroupRows(int startRow, int endRow)
    {
        _workbook.ThrowIfDisposed();
        if (startRow < 1) throw new ArgumentOutOfRangeException(nameof(startRow), startRow, "must be >= 1");
        if (endRow < 1) throw new ArgumentOutOfRangeException(nameof(endRow), endRow, "must be >= 1");
        if (startRow > endRow) throw new ArgumentOutOfRangeException(nameof(startRow), startRow, "startRow must be <= endRow");
        _underlying.UngroupRow(startRow - 1, endRow - 1);
    }

    public void GroupColumns(int startCol, int endCol)
    {
        _workbook.ThrowIfDisposed();
        if (startCol < 1) throw new ArgumentOutOfRangeException(nameof(startCol), startCol, "must be >= 1");
        if (endCol < 1) throw new ArgumentOutOfRangeException(nameof(endCol), endCol, "must be >= 1");
        if (startCol > endCol) throw new ArgumentOutOfRangeException(nameof(startCol), startCol, "startCol must be <= endCol");
        _underlying.GroupColumn(startCol - 1, endCol - 1);
    }

    public void UngroupColumns(int startCol, int endCol)
    {
        _workbook.ThrowIfDisposed();
        if (startCol < 1) throw new ArgumentOutOfRangeException(nameof(startCol), startCol, "must be >= 1");
        if (endCol < 1) throw new ArgumentOutOfRangeException(nameof(endCol), endCol, "must be >= 1");
        if (startCol > endCol) throw new ArgumentOutOfRangeException(nameof(startCol), startCol, "startCol must be <= endCol");
        _underlying.UngroupColumn(startCol - 1, endCol - 1);
    }

    public void SetRowGroupCollapsed(int row, bool collapsed)
    {
        _workbook.ThrowIfDisposed();
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row), row, "must be >= 1");
        _underlying.SetRowGroupCollapsed(row - 1, collapsed);
    }

    public void CreateSplitPane(int xSplitTwips, int ySplitTwips)
    {
        _workbook.ThrowIfDisposed();
        if (xSplitTwips < 0) throw new ArgumentOutOfRangeException(nameof(xSplitTwips), xSplitTwips, "must be >= 0");
        if (ySplitTwips < 0) throw new ArgumentOutOfRangeException(nameof(ySplitTwips), ySplitTwips, "must be >= 0");
        _underlying.CreateSplitPane(xSplitTwips, ySplitTwips, 0, 0, NPOI.SS.UserModel.PanePosition.LowerRight);
    }

    public void AddConditionalFormatting(string a1Range, params ConditionalFormat[] rules)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Range);
        ArgumentNullException.ThrowIfNull(rules);
        if (rules.Length == 0) throw new ArgumentException("At least one rule is required.", nameof(rules));

        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        var region = new NPOI.SS.Util.CellRangeAddress(r1 - 1, r2 - 1, c1 - 1, c2 - 1);
        var scf = _underlying.SheetConditionalFormatting;

        var npoiRules = new NPOI.SS.UserModel.IConditionalFormattingRule[rules.Length];
        for (int i = 0; i < rules.Length; i++)
            npoiRules[i] = rules[i].CreateNpoiRule(scf);

        scf.AddConditionalFormatting(new[] { region }, npoiRules);
    }

    public int ConditionalFormattingCount
    {
        get
        {
            _workbook.ThrowIfDisposed();
            return _underlying.SheetConditionalFormatting.NumConditionalFormattings;
        }
    }

    public void RemoveConditionalFormatting(int index)
    {
        _workbook.ThrowIfDisposed();
        _underlying.SheetConditionalFormatting.RemoveConditionalFormatting(index);
    }

    public void SortRange(string a1Range, params SortKey[] keys)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Range);
        ArgumentNullException.ThrowIfNull(keys);
        if (keys.Length == 0) throw new ArgumentException("At least one sort key is required.", nameof(keys));

        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);

        int rowCount = r2 - r1 + 1;
        if (rowCount <= 1) return;

        int colCount = c2 - c1 + 1;

        // Snapshot all cell values + styles
        var snapshot = new CellSnapshot[rowCount][];
        for (int ri = 0; ri < rowCount; ri++)
        {
            snapshot[ri] = new CellSnapshot[colCount];
            var npoiRow = _underlying.GetRow(r1 - 1 + ri);
            for (int ci = 0; ci < colCount; ci++)
            {
                var npoiCell = npoiRow?.GetCell(c1 - 1 + ci);
                snapshot[ri][ci] = CellSnapshot.Capture(npoiCell);
            }
        }

        // Sort rows by keys
        Array.Sort(snapshot, (a, b) =>
        {
            foreach (var key in keys)
            {
                int colIdx = key.Column - c1;
                if (colIdx < 0 || colIdx >= colCount) continue;
                int cmp = CellSnapshot.Compare(a[colIdx], b[colIdx]);
                if (cmp != 0) return key.Ascending ? cmp : -cmp;
            }
            return 0;
        });

        // Write sorted values back
        for (int ri = 0; ri < rowCount; ri++)
        {
            var npoiRow = _underlying.GetRow(r1 - 1 + ri) ?? _underlying.CreateRow(r1 - 1 + ri);
            for (int ci = 0; ci < colCount; ci++)
            {
                var npoiCell = npoiRow.GetCell(c1 - 1 + ci) ?? npoiRow.CreateCell(c1 - 1 + ci);
                snapshot[ri][ci].Apply(npoiCell);
            }
        }
    }

    private readonly struct CellSnapshot
    {
        public NPOI.SS.UserModel.CellType Type { get; init; }
        public double NumericValue { get; init; }
        public string? StringValue { get; init; }
        public bool BoolValue { get; init; }
        public NPOI.SS.UserModel.ICellStyle? Style { get; init; }
        public string? Formula { get; init; }

        public static CellSnapshot Capture(NPOI.SS.UserModel.ICell? cell)
        {
            if (cell == null)
                return new CellSnapshot { Type = NPOI.SS.UserModel.CellType.Blank };

            var type = cell.CellType;
            return new CellSnapshot
            {
                Type = type,
                NumericValue = type == NPOI.SS.UserModel.CellType.Numeric ? cell.NumericCellValue : 0,
                StringValue = type == NPOI.SS.UserModel.CellType.String ? cell.StringCellValue : null,
                BoolValue = type == NPOI.SS.UserModel.CellType.Boolean ? cell.BooleanCellValue : false,
                Style = cell.CellStyle,
                Formula = type == NPOI.SS.UserModel.CellType.Formula ? cell.CellFormula : null,
            };
        }

        public void Apply(NPOI.SS.UserModel.ICell cell)
        {
            switch (Type)
            {
                case NPOI.SS.UserModel.CellType.Numeric:
                    cell.SetCellValue(NumericValue);
                    break;
                case NPOI.SS.UserModel.CellType.String:
                    cell.SetCellValue(StringValue ?? "");
                    break;
                case NPOI.SS.UserModel.CellType.Boolean:
                    cell.SetCellValue(BoolValue);
                    break;
                case NPOI.SS.UserModel.CellType.Formula:
                    cell.SetCellFormula(Formula);
                    break;
                default:
                    cell.SetBlank();
                    break;
            }
            if (Style != null) cell.CellStyle = Style;
        }

        public static int Compare(CellSnapshot a, CellSnapshot b)
        {
            // Blanks sort last
            if (a.Type == NPOI.SS.UserModel.CellType.Blank && b.Type == NPOI.SS.UserModel.CellType.Blank) return 0;
            if (a.Type == NPOI.SS.UserModel.CellType.Blank) return 1;
            if (b.Type == NPOI.SS.UserModel.CellType.Blank) return -1;

            // Numeric before string (Excel behavior)
            if (a.Type == NPOI.SS.UserModel.CellType.Numeric && b.Type == NPOI.SS.UserModel.CellType.Numeric)
                return a.NumericValue.CompareTo(b.NumericValue);

            if (a.Type == NPOI.SS.UserModel.CellType.String && b.Type == NPOI.SS.UserModel.CellType.String)
                return StringComparer.OrdinalIgnoreCase.Compare(a.StringValue, b.StringValue);

            // Mixed: numbers sort before strings (Excel default)
            if (a.Type == NPOI.SS.UserModel.CellType.Numeric) return -1;
            if (b.Type == NPOI.SS.UserModel.CellType.Numeric) return 1;

            // Booleans: FALSE < TRUE
            if (a.Type == NPOI.SS.UserModel.CellType.Boolean && b.Type == NPOI.SS.UserModel.CellType.Boolean)
                return a.BoolValue.CompareTo(b.BoolValue);

            return 0;
        }
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
