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
using System.Runtime.CompilerServices;

namespace NetXlsx;

internal sealed class OoxmlSheet : ISheet
{
    private readonly OoxmlWorkbook _workbook;
    private string _name;

    internal OoxmlSheet(OoxmlWorkbook workbook, string name)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _name = name ?? throw new ArgumentNullException(nameof(name));
    }

    public string Name
    {
        get { _workbook.ThrowIfDisposed(); return _name; }
    }

    public IWorkbook Workbook
    {
        get { _workbook.ThrowIfDisposed(); return _workbook; }
    }

    // Allows the workbook to keep wrapper names in sync if a rename API lands.
    internal void SetNameInternal(string name) => _name = name;

    // ---- Not-yet-implemented surface (lands slice by slice; see I-82) -------

    private static NotImplementedException NotYet([CallerMemberName] string? member = null)
        => new(
            $"ISheet.{member} is not yet implemented on the Open XML SDK engine " +
            "(I-82 engine swap). It lands in a later slice; until then use the " +
            "legacy engine (Workbook.Create/Open) for this operation, or track " +
            "the swap in docs/design.md (I-82).");

    public ICell this[string a1] => throw NotYet();
    public ICell this[int row, int column] => throw NotYet();

    public IRange Range(string a1Range) => throw NotYet();
    public IRange Range(int row1, int col1, int row2, int col2) => throw NotYet();

    public IRow AppendRow() => throw NotYet();
    public IRow Row(int index) => throw NotYet();

    public IColumn Column(int index) => throw NotYet();
    public IColumn Column(string letter) => throw NotYet();

    public void FreezeRows(int rows) => throw NotYet();
    public void FreezeColumns(int cols) => throw NotYet();
    public void FreezePane(int rows, int cols) => throw NotYet();

    public void GroupRows(int startRow, int endRow) => throw NotYet();
    public void UngroupRows(int startRow, int endRow) => throw NotYet();
    public void GroupColumns(int startCol, int endCol) => throw NotYet();
    public void UngroupColumns(int startCol, int endCol) => throw NotYet();
    public void SetRowGroupCollapsed(int row, bool collapsed) => throw NotYet();
    public void CreateSplitPane(int xSplitTwips, int ySplitTwips) => throw NotYet();

    public IChart AddChart(ChartType type, string startCell, string endCell, string categoryRange, string valueRange, string? title = null) => throw NotYet();
    public IShape AddShape(ShapeType type, string startCell, string endCell, Color? fillColor = null, Color? lineColor = null) => throw NotYet();
    public IConnector AddConnector(ConnectorType type, string startCell, string endCell,
        Color? lineColor = null,
        int dx1 = 0, int dy1 = 0, int dx2 = 0, int dy2 = 0,
        bool flipH = false, bool flipV = false,
        ConnectorEnd headEnd = ConnectorEnd.None, ConnectorEnd tailEnd = ConnectorEnd.None,
        double? lineWidthPoints = null) => throw NotYet();

    public IReadOnlyList<IPicture> Pictures => throw NotYet();
    public IReadOnlyList<IConnector> Connectors => throw NotYet();

    public void AddConditionalFormatting(string a1Range, params ConditionalFormat[] rules) => throw NotYet();
    public int ConditionalFormattingCount => throw NotYet();
    public void RemoveConditionalFormatting(int index) => throw NotYet();

    public void SortRange(string a1Range, params SortKey[] keys) => throw NotYet();

    public void MergeCells(string a1Range) => throw NotYet();
    public void MergeCellsStyled(string a1Range, CellStyle style) => throw NotYet();
    public void UnmergeCells(string a1Range) => throw NotYet();
    public IReadOnlyList<string> MergedRanges => throw NotYet();

    public bool Hidden { get => throw NotYet(); set => throw NotYet(); }
    public bool ShowGridlines { get => throw NotYet(); set => throw NotYet(); }
    public double? DefaultColumnWidth { get => throw NotYet(); set => throw NotYet(); }

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

    public IPicture AddPicture(string a1Cell, byte[] data, ImageFormat format) => throw NotYet();
    public IPicture AddPicture(string a1Cell, byte[] data) => throw NotYet();
    public IPicture AddPicture(string startCell, string endCell, byte[] data, ImageFormat format) => throw NotYet();
    public IPicture AddPicture(string startCell, string endCell, byte[] data) => throw NotYet();
    public IPicture AddPicture(string startCell, string endCell, byte[] data, ImageFormat format,
        int dx1, int dy1, int dx2, int dy2) => throw NotYet();

    public void Protect(string? password = null, SheetProtection? options = null) => throw NotYet();
    public void Unprotect() => throw NotYet();
    public bool IsProtected => throw NotYet();

    // Escape hatch divergence (I-82): no NPOI sheet exists on the SDK engine.
    public NPOI.XSSF.UserModel.XSSFSheet Underlying => throw new NotSupportedException(
        "ISheet.Underlying (NPOI XSSFSheet) is not available on the Open XML SDK " +
        "engine (I-82). Use IWorkbook.OpenXmlDocument for the SDK escape hatch.");
}
