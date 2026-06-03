// I-82 engine swap — streaming row + the buffered row/cell data model
// (slice 9). A row lives in its sheet's bounded window until flushed through
// the forward-only writer; the buffer types here are the in-window
// representation (plain values, not DOM nodes).

using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlStreamingRow : IStreamingRow
{
    private readonly OoxmlStreamingWorkbook _workbook;
    private readonly OoxmlStreamingSheet _sheet;
    private readonly StreamingRowBuffer _buffer;

    internal OoxmlStreamingRow(OoxmlStreamingWorkbook workbook, OoxmlStreamingSheet sheet, StreamingRowBuffer buffer)
    {
        _workbook = workbook;
        _sheet = sheet;
        _buffer = buffer;
    }

    public int Index
    {
        get { _workbook.ThrowIfDisposed(); return _buffer.Row1; }
    }

    public IStreamingSheet Sheet
    {
        get { _workbook.ThrowIfDisposed(); return _sheet; }
    }

    public IStreamingCell Cell(int column)
    {
        _workbook.ThrowIfDisposed();
        if (column < 1 || column > CellAddress.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column,
                $"column must be in [1, {CellAddress.MaxColumn}]");
        return new OoxmlStreamingCell(_workbook, _buffer, column);
    }

    public IStreamingCell this[int column] => Cell(column);

    public IStreamingCell this[string columnLetter]
    {
        get
        {
            _workbook.ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(columnLetter);
            int col = CellAddress.ParseColumn(columnLetter);
            return new OoxmlStreamingCell(_workbook, _buffer, col);
        }
    }

    public IStreamingRow Set(int column, string value) { Cell(column).SetString(value); return this; }
    public IStreamingRow Set(int column, double value) { Cell(column).SetNumber(value); return this; }
    public IStreamingRow Set(int column, decimal value) { Cell(column).SetNumber(value); return this; }
    public IStreamingRow Set(int column, int value) { Cell(column).SetNumber(value); return this; }
    public IStreamingRow Set(int column, long value) { Cell(column).SetNumber(value); return this; }
    public IStreamingRow Set(int column, bool value) { Cell(column).SetBool(value); return this; }
    public IStreamingRow Set(int column, DateTime value) { Cell(column).SetDate(value); return this; }

    public void Flush()
    {
        _workbook.ThrowIfDisposed();
        // SXSSFSheet.FlushRows() semantics: every buffered row (this one
        // included) goes to disk; the window empties.
        _sheet.FlushAllBuffered();
    }
}

/// <summary>
/// One in-window row: its 1-based index and the cells written so far. Flushing
/// releases the cell data — a flushed row is on disk and cannot be revisited.
/// </summary>
internal sealed class StreamingRowBuffer
{
    internal StreamingRowBuffer(int row1) => Row1 = row1;

    internal int Row1 { get; }
    internal SortedDictionary<int, StreamingCellData> Cells { get; } = new();
    internal bool Flushed { get; private set; }

    internal void MarkFlushed()
    {
        Flushed = true;
        Cells.Clear(); // drop the data — bounded memory is the contract
    }
}

/// <summary>
/// One buffered cell value. Materialized into the OOXML &lt;c&gt; shape only at
/// flush time; the shapes mirror the random-access SDK engine exactly (inline
/// strings with space preservation, G17 invariant numbers, 1/0 booleans, no
/// cached &lt;v&gt; on formulas per decision #46).
/// </summary>
internal sealed class StreamingCellData
{
    internal CellKind Kind = CellKind.Empty; // Date is derived from StyleIdx, never stored
    internal string? Str;
    internal double Num;
    internal bool Bool;
    internal string? FormulaBody;
    internal uint StyleIdx;
    internal CellStyle? AppliedStyle; // last NetXlsx-applied style (merge fallback)

    internal S.Cell ToCell(string cellRef)
    {
        var c = new S.Cell { CellReference = cellRef };
        if (StyleIdx != 0) c.StyleIndex = StyleIdx;
        switch (Kind)
        {
            case CellKind.String:
                c.DataType = S.CellValues.InlineString;
                c.AppendChild(new S.InlineString(new S.Text(Str!) { Space = SpaceProcessingModeValues.Preserve }));
                break;
            case CellKind.Number:
                // Numeric is the default type — no t attribute, like the
                // random-access SDK engine (NPOI streaming writes t="n";
                // both forms are equivalent on read).
                c.CellValue = new S.CellValue(Num.ToString("G17", CultureInfo.InvariantCulture));
                break;
            case CellKind.Bool:
                c.DataType = S.CellValues.Boolean;
                c.CellValue = new S.CellValue(Bool ? "1" : "0");
                break;
            case CellKind.Formula:
                // No cached <v> (decision #46 / §7.8): Excel recalculates on open.
                c.CellFormula = new S.CellFormula(FormulaBody!);
                break;
            case CellKind.Empty:
            default:
                break; // styled-but-valueless cell still persists as a style carrier
        }
        return c;
    }
}
