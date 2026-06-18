// I-92: placeholder ISheet for a chartsheet / dialogsheet opened from a file.
//
// A chartsheet (full-window chart, no cell grid) and the legacy dialogsheet
// are legal, Excel-authorable sheet kinds that appear in workbook tab order.
// NetXlsx does not model their content; it opens them as placeholders so the
// whole workbook is usable and round-trips byte-stable (the backing part is
// preserved verbatim by the clone-Save — it is reachable in the relationship
// graph). The sheet participates in SheetCount / Sheets / the indexers, carries
// Name / Hidden, and takes part in rename / move / remove; every cell-/grid-
// shaped member throws NotSupportedException (use IWorkbook.Underlying to reach
// the underlying ChartsheetPart for advanced work).

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlChartsheet : IOoxmlSheet
{
    private readonly OoxmlWorkbook _workbook;
    private readonly OpenXmlPart _part;     // ChartsheetPart or DialogsheetPart
    private readonly SheetKind _kind;
    private string _name;
    private bool _removed;

    internal OoxmlChartsheet(OoxmlWorkbook workbook, string name, OpenXmlPart part, SheetKind kind)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _name = name ?? throw new ArgumentNullException(nameof(name));
        _part = part ?? throw new ArgumentNullException(nameof(part));
        _kind = kind;
    }

    // ---- Supported surface --------------------------------------------------

    public string Name { get { ThrowIfUnusable(); return _name; } }

    public IWorkbook Workbook { get { ThrowIfUnusable(); return _workbook; } }

    public SheetKind Kind { get { ThrowIfUnusable(); return _kind; } }

    public void Rename(string newName)
    {
        ThrowIfUnusable();
        _workbook.RenameSheet(this, newName);
    }

    public bool Hidden
    {
        get { ThrowIfUnusable(); return IsHiddenInternal; }
        set
        {
            ThrowIfUnusable();
            var sheet = _workbook.SheetElementFor(_part);
            sheet.State = value ? S.SheetStateValues.Hidden : null;
        }
    }

    // ---- IOoxmlSheet (workbook-side lifecycle) ------------------------------

    public OoxmlWorkbook WorkbookInternal => _workbook;

    public OpenXmlPart SheetPartInternal => _part;

    public bool IsHiddenInternal
    {
        get
        {
            var state = _workbook.SheetElementFor(_part).State?.Value;
            return state == S.SheetStateValues.Hidden || state == S.SheetStateValues.VeryHidden;
        }
    }

    public void SetNameInternal(string name) => _name = name;

    public void MarkRemoved() => _removed = true;

    // A chartsheet/dialogsheet has no formula-shaped surface, so it never holds
    // references to other sheets — both rewrite passes are no-ops. (References
    // *to* this sheet from other worksheets are rewritten by the workbook's
    // fan-out loop, which walks every sheet regardless of kind.)
    public void RewriteSheetReferences(string oldName, string newName) { }

    public void RewriteSheetReferencesToRefError(string removedName) { }

    // Mirrors OoxmlSheet.ThrowIfUnusable: disposed first (ObjectDisposedException),
    // then removed (InvalidOperationException). The throwing grid members below
    // need no usability check — they are unsupported in every state.
    private void ThrowIfUnusable()
    {
        _workbook.ThrowIfDisposed();
        if (_removed)
            throw new InvalidOperationException($"sheet '{_name}' has been removed from the workbook.");
    }

    private NotSupportedException NotSupported(string member)
        => new($"'{member}' is not supported on sheet '{_name}' because it is a " +
               $"{_kind.ToString().ToLowerInvariant()} with no cell grid. " +
               "Use IWorkbook.Underlying to reach the underlying part.");

    // ---- Unsupported surface (no cell grid) ---------------------------------

    public ICell this[string a1] => throw NotSupported("this[string]");
    public ICell this[int row, int column] => throw NotSupported("this[int, int]");
    public IRange Range(string a1Range) => throw NotSupported(nameof(Range));
    public IRange Range(int row1, int col1, int row2, int col2) => throw NotSupported(nameof(Range));
    public IRow AppendRow() => throw NotSupported(nameof(AppendRow));
    public IRow Row(int index) => throw NotSupported(nameof(Row));
    public int LastRowNumber => throw NotSupported(nameof(LastRowNumber));
    public IColumn Column(int index) => throw NotSupported(nameof(Column));
    public IColumn Column(string letter) => throw NotSupported(nameof(Column));
    public void FreezeRows(int rows) => throw NotSupported(nameof(FreezeRows));
    public void FreezeColumns(int cols) => throw NotSupported(nameof(FreezeColumns));
    public void FreezePane(int rows, int cols) => throw NotSupported(nameof(FreezePane));
    public void GroupRows(int startRow, int endRow) => throw NotSupported(nameof(GroupRows));
    public void UngroupRows(int startRow, int endRow) => throw NotSupported(nameof(UngroupRows));
    public void GroupColumns(int startCol, int endCol) => throw NotSupported(nameof(GroupColumns));
    public void UngroupColumns(int startCol, int endCol) => throw NotSupported(nameof(UngroupColumns));
    public void SetRowGroupCollapsed(int row, bool collapsed) => throw NotSupported(nameof(SetRowGroupCollapsed));
    public void CreateSplitPane(int xSplitTwips, int ySplitTwips) => throw NotSupported(nameof(CreateSplitPane));

    public IChart AddChart(ChartType type, string startCell, string endCell, string categoryRange, string valueRange, string? title = null)
        => throw NotSupported(nameof(AddChart));
    public IShape AddShape(ShapeType type, string startCell, string endCell, Color? fillColor = null, Color? lineColor = null)
        => throw NotSupported(nameof(AddShape));
    public IConnector AddConnector(ConnectorType type, string startCell, string endCell,
        Color? lineColor = null,
        int dx1 = 0, int dy1 = 0, int dx2 = 0, int dy2 = 0,
        bool flipH = false, bool flipV = false,
        ConnectorEnd headEnd = ConnectorEnd.None, ConnectorEnd tailEnd = ConnectorEnd.None,
        double? lineWidthPoints = null)
        => throw NotSupported(nameof(AddConnector));
    public IReadOnlyList<IPicture> Pictures => throw NotSupported(nameof(Pictures));
    public IReadOnlyList<IConnector> Connectors => throw NotSupported(nameof(Connectors));
    public void RemovePicture(IPicture picture) => throw NotSupported(nameof(RemovePicture));
    public void RemoveChart(IChart chart) => throw NotSupported(nameof(RemoveChart));
    public void RemoveShape(IShape shape) => throw NotSupported(nameof(RemoveShape));
    public void RemoveConnector(IConnector connector) => throw NotSupported(nameof(RemoveConnector));

    public void AddConditionalFormatting(string a1Range, params ConditionalFormat[] rules)
        => throw NotSupported(nameof(AddConditionalFormatting));
    public int ConditionalFormattingCount => throw NotSupported(nameof(ConditionalFormattingCount));
    public void RemoveConditionalFormatting(int index) => throw NotSupported(nameof(RemoveConditionalFormatting));

    public void SortRange(string a1Range, params SortKey[] keys) => throw NotSupported(nameof(SortRange));

    public void MergeCells(string a1Range) => throw NotSupported(nameof(MergeCells));
    public void MergeCellsStyled(string a1Range, CellStyle style) => throw NotSupported(nameof(MergeCellsStyled));
    public void UnmergeCells(string a1Range) => throw NotSupported(nameof(UnmergeCells));
    public IReadOnlyList<string> MergedRanges => throw NotSupported(nameof(MergedRanges));

    public bool ShowGridlines
    {
        get => throw NotSupported(nameof(ShowGridlines));
        set => throw NotSupported(nameof(ShowGridlines));
    }
    public double? DefaultColumnWidth
    {
        get => throw NotSupported(nameof(DefaultColumnWidth));
        set => throw NotSupported(nameof(DefaultColumnWidth));
    }

    public ITable AddTable(string a1Range, string name, string? style = null) => throw NotSupported(nameof(AddTable));
    public IReadOnlyList<ITable> Tables => throw NotSupported(nameof(Tables));
    public bool TryGetTable(string name, [MaybeNullWhen(false)] out ITable table) => throw NotSupported(nameof(TryGetTable));
    public void RemoveTable(ITable table) => throw NotSupported(nameof(RemoveTable));

    public void SetAutoFilter(string a1Range) => throw NotSupported(nameof(SetAutoFilter));
    public void ClearAutoFilter() => throw NotSupported(nameof(ClearAutoFilter));
    public void SetAutoFilterColumn(int columnOffset, FilterCriteria criteria) => throw NotSupported(nameof(SetAutoFilterColumn));
    public void ClearAutoFilterColumn(int columnOffset) => throw NotSupported(nameof(ClearAutoFilterColumn));
    public bool HasAutoFilter => throw NotSupported(nameof(HasAutoFilter));
    public string? AutoFilterRange => throw NotSupported(nameof(AutoFilterRange));

    public void AddValidation(string a1Range, DataValidation validation) => throw NotSupported(nameof(AddValidation));
    public void RemoveValidation(string a1Range) => throw NotSupported(nameof(RemoveValidation));

    public IPicture AddPicture(string a1Cell, byte[] data, ImageFormat format) => throw NotSupported(nameof(AddPicture));
    public IPicture AddPicture(string a1Cell, byte[] data) => throw NotSupported(nameof(AddPicture));
    public IPicture AddPicture(string startCell, string endCell, byte[] data, ImageFormat format) => throw NotSupported(nameof(AddPicture));
    public IPicture AddPicture(string startCell, string endCell, byte[] data) => throw NotSupported(nameof(AddPicture));
    public IPicture AddPicture(string startCell, string endCell, byte[] data, ImageFormat format, int dx1, int dy1, int dx2, int dy2)
        => throw NotSupported(nameof(AddPicture));

    public void Protect(string? password = null, SheetProtection? options = null) => throw NotSupported(nameof(Protect));
    public void Unprotect() => throw NotSupported(nameof(Unprotect));
    public bool IsProtected => throw NotSupported(nameof(IsProtected));

    public S.Worksheet Underlying => throw NotSupported(nameof(Underlying));
}
