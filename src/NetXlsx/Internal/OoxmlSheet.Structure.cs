// I-82 engine swap — Open XML SDK-backed sheet structure (structure slice, 5b).
//
// Frozen/split panes, row/column grouping (outline), sheet visibility, gridlines,
// and default column width. Each mirrors the NPOI engine's contract (XssfSheet) on
// the SDK engine, writing the OOXML node directly:
//
//   * panes      -> <sheetView><pane>/<selection>           (xSplit/ySplit/state)
//   * grouping   -> <row outlineLevel>/<col outlineLevel> + <sheetFormatPr>
//   * visibility -> workbook.xml <sheet @state>             (NOT on the worksheet)
//   * gridlines  -> <sheetView @showGridLines>              (default true)
//   * defColWidth-> <sheetFormatPr @defaultColWidth>        (lesson #3)
//
// All new worksheet children are placed by OoxmlSchemaOrder so they land in
// CT_Worksheet's strict child order even on opened files (SDK-quirk #3 / #8).
// <sheetFormatPr> carries the required @defaultRowHeight whenever it is created.

using System;
using System.Linq;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    // Excel's default row height in points — the required CT_SheetFormatPr
    // @defaultRowHeight whenever the engine has to materialize <sheetFormatPr>.
    private const double DefaultRowHeightPoints = 15D;

    // ---- Frozen / split panes ----------------------------------------------
    // <sheetView><pane>: xSplit = frozen columns, ySplit = frozen rows,
    // topLeftCell = first unfrozen cell, activePane + a <selection> in that pane,
    // state = frozen (panes) or split (draggable). Mirrors NPOI's
    // CreateFreezePane(colSplit, rowSplit) — column first — and CreateSplitPane.

    public void FreezeRows(int rows)
    {
        ThrowIfUnusable();
        FreezePane(rows, 0);
    }

    public void FreezeColumns(int cols)
    {
        ThrowIfUnusable();
        FreezePane(0, cols);
    }

    public void FreezePane(int rows, int cols)
    {
        ThrowIfUnusable();
        if (rows < 0) throw new ArgumentOutOfRangeException(nameof(rows), rows, "must be >= 0");
        if (cols < 0) throw new ArgumentOutOfRangeException(nameof(cols), cols, "must be >= 0");

        var view = GetOrCreateDefaultSheetView();
        ClearSelections(view);

        // (0, 0) clears any existing freeze.
        if (rows == 0 && cols == 0)
        {
            view.GetFirstChild<S.Pane>()?.Remove();
            return;
        }

        var pane = GetOrCreatePane(view);
        pane.HorizontalSplit = cols > 0 ? cols : null;   // xSplit (columns)
        pane.VerticalSplit = rows > 0 ? rows : null;     // ySplit (rows)
        pane.State = S.PaneStateValues.Frozen;

        string topLeft;
        S.PaneValues active;
        if (rows == 0)
        {
            topLeft = CellAddress.Format(1, cols + 1);
            active = S.PaneValues.TopRight;
        }
        else if (cols == 0)
        {
            topLeft = CellAddress.Format(rows + 1, 1);
            active = S.PaneValues.BottomLeft;
        }
        else
        {
            topLeft = CellAddress.Format(rows + 1, cols + 1);
            active = S.PaneValues.BottomRight;
        }
        pane.TopLeftCell = topLeft;
        pane.ActivePane = active;
        view.InsertAfter(new S.Selection { Pane = active }, pane);
    }

    public void CreateSplitPane(int xSplitTwips, int ySplitTwips)
    {
        ThrowIfUnusable();
        if (xSplitTwips < 0) throw new ArgumentOutOfRangeException(nameof(xSplitTwips), xSplitTwips, "must be >= 0");
        if (ySplitTwips < 0) throw new ArgumentOutOfRangeException(nameof(ySplitTwips), ySplitTwips, "must be >= 0");

        var view = GetOrCreateDefaultSheetView();
        ClearSelections(view);

        if (xSplitTwips == 0 && ySplitTwips == 0)
        {
            view.GetFirstChild<S.Pane>()?.Remove();
            return;
        }

        var pane = GetOrCreatePane(view);
        pane.HorizontalSplit = xSplitTwips > 0 ? xSplitTwips : null;
        pane.VerticalSplit = ySplitTwips > 0 ? ySplitTwips : null;
        // A split pane is draggable (state=split) with the lower-right pane active,
        // mirroring NPOI's CreateSplitPane(..., PanePosition.LowerRight).
        pane.TopLeftCell = "A1";
        pane.State = S.PaneStateValues.Split;
        pane.ActivePane = S.PaneValues.BottomRight;
        view.InsertAfter(new S.Selection { Pane = S.PaneValues.BottomRight }, pane);
    }

    // ---- Gridlines ----------------------------------------------------------

    public bool ShowGridlines
    {
        get
        {
            ThrowIfUnusable();
            // Absent @showGridLines means shown — Excel's default is true.
            return FindDefaultSheetView()?.ShowGridLines?.Value ?? true;
        }
        set
        {
            ThrowIfUnusable();
            var view = GetOrCreateDefaultSheetView();
            // true is the default — clear the attribute rather than emit noise.
            view.ShowGridLines = value ? null : false;
        }
    }

    // ---- Sheet visibility (workbook.xml <sheet @state>) ---------------------

    public bool Hidden
    {
        get
        {
            ThrowIfUnusable();
            return IsHiddenInternal;
        }
        set
        {
            ThrowIfUnusable();
            var sheet = _workbook.SheetElementFor(_worksheetPart);
            // Visible is the default — clear the attribute when un-hiding.
            sheet.State = value ? S.SheetStateValues.Hidden : null;
        }
    }

    // Non-throwing visibility read for workbook-side lifecycle code (the
    // last-visible-sheet guard in RemoveSheet counts visible sheets without
    // tripping any disposal/removed guard). Mirrors NPOI: both Hidden and
    // VeryHidden read as hidden.
    internal bool IsHiddenInternal
    {
        get
        {
            var state = _workbook.SheetElementFor(_worksheetPart).State?.Value;
            return state == S.SheetStateValues.Hidden || state == S.SheetStateValues.VeryHidden;
        }
    }

    // ---- Default column width (lesson #3) -----------------------------------

    public double? DefaultColumnWidth
    {
        get
        {
            ThrowIfUnusable();
            var w = Worksheet.GetFirstChild<S.SheetFormatProperties>()?.DefaultColumnWidth?.Value;
            // NPOI treats an absent/zero width as "unset" (Excel derives from fonts).
            return w is null or 0 ? null : w;
        }
        set
        {
            ThrowIfUnusable();
            // null (or a non-positive value) clears the attribute so Excel derives
            // the width from the Normal style's font metrics (lesson #3 / I-78).
            if (value is null || value.Value <= 0)
            {
                var existing = Worksheet.GetFirstChild<S.SheetFormatProperties>();
                if (existing is not null) existing.DefaultColumnWidth = null;
                return;
            }
            GetOrCreateSheetFormatProperties().DefaultColumnWidth = value.Value;
        }
    }

    // ---- Row grouping (outline) ---------------------------------------------
    // <row @outlineLevel> per grouped row; <sheetFormatPr @outlineLevelRow> tracks
    // the deepest level. 1-based, validated >= 1 and start <= end (mirrors NPOI).

    public void GroupRows(int startRow, int endRow)
    {
        ThrowIfUnusable();
        ValidateRange(startRow, endRow, nameof(startRow), nameof(endRow));
        for (int r = startRow; r <= endRow; r++)
        {
            var row = GetOrCreateRow(r);
            row.OutlineLevel = (byte)((row.OutlineLevel?.Value ?? 0) + 1);
        }
        UpdateRowOutlineLevel();
    }

    public void UngroupRows(int startRow, int endRow)
    {
        ThrowIfUnusable();
        ValidateRange(startRow, endRow, nameof(startRow), nameof(endRow));
        for (int r = startRow; r <= endRow; r++)
        {
            var row = FindRow(r);
            if (row is null) continue;
            byte level = row.OutlineLevel?.Value ?? 0;
            if (level > 0) level--;
            row.OutlineLevel = level == 0 ? null : level;
        }
        UpdateRowOutlineLevel();
    }

    public void GroupColumns(int startCol, int endCol)
    {
        ThrowIfUnusable();
        ValidateRange(startCol, endCol, nameof(startCol), nameof(endCol));
        for (int c = startCol; c <= endCol; c++)
        {
            var col = GetOrCreateColumn(c);
            col.OutlineLevel = (byte)((col.OutlineLevel?.Value ?? 0) + 1);
        }
        UpdateColumnOutlineLevel();
    }

    public void UngroupColumns(int startCol, int endCol)
    {
        ThrowIfUnusable();
        ValidateRange(startCol, endCol, nameof(startCol), nameof(endCol));
        for (int c = startCol; c <= endCol; c++)
        {
            var col = FindColumn(c);
            if (col is null) continue;
            byte level = col.OutlineLevel?.Value ?? 0;
            if (level > 0) level--;
            col.OutlineLevel = level == 0 ? null : level;
        }
        UpdateColumnOutlineLevel();
    }

    public void SetRowGroupCollapsed(int row, bool collapsed)
    {
        ThrowIfUnusable();
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row), row, "must be >= 1");
        if (collapsed) CollapseRowGroup(row);
        else ExpandRowGroup(row);
    }

    // Collapses the outline group containing the summary at `row`: hides every row
    // in the contiguous >= summary-level block and marks the boundary row collapsed
    // — reproducing NPOI/POI's observable behavior (the cross-checked
    // GroupingTests.SetRowGroupCollapsed_Hides_Grouped_Rows contract).
    private void CollapseRowGroup(int row)
    {
        var summary = FindRow(row);
        if (summary is null) return;
        byte level = summary.OutlineLevel?.Value ?? 0;
        int start = FindStartOfRowGroup(row, level);
        int boundary = WriteRowsHidden(start, level, hidden: true);
        GetOrCreateRow(boundary).Collapsed = true;
    }

    private void ExpandRowGroup(int row)
    {
        var summary = FindRow(row);
        if (summary is null) return;
        byte level = summary.OutlineLevel?.Value ?? 0;
        int start = FindStartOfRowGroup(row, level);
        int boundary = WriteRowsHidden(start, level, hidden: false);
        var boundaryRow = FindRow(boundary);
        if (boundaryRow is not null) boundaryRow.Collapsed = null;
    }

    // Walks up from `row` while rows stay at >= `level`; returns the first row of
    // the contiguous group block.
    private int FindStartOfRowGroup(int row, byte level)
    {
        int current = row;
        while (current >= 1)
        {
            var r = FindRow(current);
            if (r is null || (r.OutlineLevel?.Value ?? 0) < level) return current + 1;
            current--;
        }
        return current + 1;
    }

    // Sets @hidden on the contiguous block of rows at >= `level` starting at
    // `startRow`; returns the index just past the block (the boundary row).
    private int WriteRowsHidden(int startRow, byte level, bool hidden)
    {
        int idx = startRow;
        while (true)
        {
            var r = FindRow(idx);
            if (r is null || (r.OutlineLevel?.Value ?? 0) < level) break;
            r.Hidden = hidden ? true : null;
            idx++;
        }
        return idx;
    }

    private void UpdateRowOutlineLevel()
    {
        byte max = 0;
        foreach (var row in Data.Elements<S.Row>())
        {
            byte lvl = row.OutlineLevel?.Value ?? 0;
            if (lvl > max) max = lvl;
        }
        if (max == 0)
        {
            var sfp = Worksheet.GetFirstChild<S.SheetFormatProperties>();
            if (sfp is not null) sfp.OutlineLevelRow = null;
            return;
        }
        GetOrCreateSheetFormatProperties().OutlineLevelRow = max;
    }

    private void UpdateColumnOutlineLevel()
    {
        byte max = 0;
        var cols = Worksheet.GetFirstChild<S.Columns>();
        if (cols is not null)
            foreach (var col in cols.Elements<S.Column>())
            {
                byte lvl = col.OutlineLevel?.Value ?? 0;
                if (lvl > max) max = lvl;
            }
        if (max == 0)
        {
            var sfp = Worksheet.GetFirstChild<S.SheetFormatProperties>();
            if (sfp is not null) sfp.OutlineLevelColumn = null;
            return;
        }
        GetOrCreateSheetFormatProperties().OutlineLevelColumn = max;
    }

    // ---- Schema-ordered element helpers shared by the structure surface -----

    private S.SheetView? FindDefaultSheetView()
        => Worksheet.GetFirstChild<S.SheetViews>()?.GetFirstChild<S.SheetView>();

    private S.SheetView GetOrCreateDefaultSheetView()
    {
        var views = OoxmlSchemaOrder.GetOrInsert(Worksheet, static () => new S.SheetViews());
        var view = views.GetFirstChild<S.SheetView>();
        if (view is null)
        {
            // CT_SheetView requires @workbookViewId (default view 0); CT_SheetViews
            // requires >= 1 <sheetView>, so create them together.
            view = new S.SheetView { WorkbookViewId = 0U };
            views.AppendChild(view);
        }
        return view;
    }

    // <pane> is the first child of CT_SheetView, before <selection>.
    private static S.Pane GetOrCreatePane(S.SheetView view)
    {
        var pane = view.GetFirstChild<S.Pane>();
        if (pane is not null) return pane;
        pane = new S.Pane();
        view.InsertAt(pane, 0);
        return pane;
    }

    private static void ClearSelections(S.SheetView view)
    {
        foreach (var sel in view.Elements<S.Selection>().ToList()) sel.Remove();
    }

    // CT_SheetFormatPr requires @defaultRowHeight, so set it whenever the engine
    // first materializes the element.
    private S.SheetFormatProperties GetOrCreateSheetFormatProperties()
        => OoxmlSchemaOrder.GetOrInsert(Worksheet,
            static () => new S.SheetFormatProperties { DefaultRowHeight = DefaultRowHeightPoints });

    private static void ValidateRange(int start, int end, string startName, string endName)
    {
        if (start < 1) throw new ArgumentOutOfRangeException(startName, start, "must be >= 1");
        if (end < 1) throw new ArgumentOutOfRangeException(endName, end, "must be >= 1");
        if (start > end) throw new ArgumentOutOfRangeException(startName, start, $"{startName} must be <= {endName}");
    }
}
