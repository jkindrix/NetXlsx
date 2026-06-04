using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace NetXlsx;

/// <summary>
/// A write-only, append-only workbook backed by a forward-only
/// streaming writer with bounded memory. Deliberately narrower than
/// <see cref="IWorkbook"/>: random access is absent because once a row
/// is flushed past the window, it cannot be revisited (design §6.3,
/// decision #7).
/// <para>
/// Use via <see cref="Workbook.CreateStreaming"/>. For random-access
/// workbooks (read, edit, multi-pass write) use
/// <see cref="Workbook.Create"/> instead.
/// </para>
/// </summary>
public interface IStreamingWorkbook : IDisposable, IAsyncDisposable
{
    /// <summary>Adds a new sheet to the workbook.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="name"/> is null.</exception>
    /// <exception cref="SheetNameException">The name violates Excel's sheet-naming rules or duplicates an existing sheet (case-insensitive).</exception>
    IStreamingSheet AddSheet(string name);

    /// <summary>Saves the workbook synchronously to <paramref name="path"/>.</summary>
    void Save(string path);

    /// <summary>Saves the workbook synchronously to <paramref name="stream"/>.</summary>
    void Save(Stream stream, bool leaveOpen = true);

    /// <summary>Saves the workbook asynchronously to <paramref name="path"/>.</summary>
    Task SaveAsync(string path, CancellationToken ct = default);

    /// <summary>Saves the workbook asynchronously to <paramref name="stream"/>.</summary>
    Task SaveAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default);

    // No escape hatch (v2.0.0 / I-82): the streaming engine writes rows
    // forward-only through OpenXmlWriter into per-sheet temp streams and
    // assembles the package only at Save — there is no live document
    // object to expose at any point before then.
}

/// <summary>
/// A sheet on an <see cref="IStreamingWorkbook"/>. Supports append-only
/// row writes; random-access row APIs are absent by design.
/// </summary>
public interface IStreamingSheet
{
    /// <summary>The sheet's name.</summary>
    string Name { get; }

    /// <summary>The owning streaming workbook.</summary>
    IStreamingWorkbook Workbook { get; }

    /// <summary>
    /// Appends a new row at the next index after the most recently
    /// written row. For an empty sheet, creates row 1.
    /// </summary>
    /// <exception cref="InvalidOperationException">Appending would exceed Excel's row limit.</exception>
    IStreamingRow AppendRow();

    /// <summary>
    /// Appends a row at an explicit 1-based <paramref name="index"/>.
    /// Must be strictly greater than the last index written; once a
    /// row is flushed out of the access window, earlier indices
    /// cannot be revisited.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is &lt;= last written index or outside Excel's grid.</exception>
    IStreamingRow AppendRow(int index);

    // No escape hatch (v2.0.0 / I-82) — see IStreamingWorkbook.
}

/// <summary>
/// A row on an <see cref="IStreamingSheet"/>. Cells are exposed via
/// the narrower <see cref="IStreamingCell"/> interface — full
/// <see cref="ICell"/> would have to honor a DOM-typed <c>Underlying</c>
/// escape hatch, which a streaming engine cannot provide: buffered rows
/// have no per-cell DOM node to expose before they are flushed to the
/// writer.
/// </summary>
public interface IStreamingRow
{
    /// <summary>The 1-based row index.</summary>
    int Index { get; }

    /// <summary>The owning sheet.</summary>
    IStreamingSheet Sheet { get; }

    /// <summary>Looks up a cell by 1-based column index.</summary>
    /// <exception cref="ArgumentOutOfRangeException">Column out of Excel's grid.</exception>
    IStreamingCell Cell(int column);

    /// <summary>Indexer form of <see cref="Cell(int)"/>.</summary>
    IStreamingCell this[int column] { get; }

    /// <summary>Indexer keyed by column letter (e.g. <c>"A"</c>).</summary>
    /// <exception cref="InvalidCellAddressException">Not a valid column letter.</exception>
    IStreamingCell this[string columnLetter] { get; }

    /// <summary>
    /// Writes a string value to the column. Returns this row for
    /// chaining (mirrors <see cref="IRow.Set(int, string)"/>).
    /// </summary>
    IStreamingRow Set(int column, string value);
    /// <summary>Writes a double value.</summary>
    IStreamingRow Set(int column, double value);
    /// <summary>Writes a decimal value.</summary>
    IStreamingRow Set(int column, decimal value);
    /// <summary>Writes an int value.</summary>
    IStreamingRow Set(int column, int value);
    /// <summary>Writes a long value.</summary>
    IStreamingRow Set(int column, long value);
    /// <summary>Writes a bool value.</summary>
    IStreamingRow Set(int column, bool value);
    /// <summary>Writes a <see cref="DateTime"/> value.</summary>
    IStreamingRow Set(int column, DateTime value);

    /// <summary>
    /// Flushes any cached rows preceding this one to disk. Called
    /// implicitly by the next <see cref="IStreamingSheet.AppendRow()"/>
    /// once the row-access window is exceeded; call explicitly to
    /// release memory at a specific point.
    /// </summary>
    void Flush();
}

/// <summary>
/// A cell on an <see cref="IStreamingRow"/>. Provides value setters,
/// styling, and the cell-kind getter — but no <c>Underlying</c> escape
/// hatch: rows are serialized forward-only through <c>OpenXmlWriter</c>
/// and the package is assembled only at <c>Save</c>, so there is no live
/// document object to return at any earlier point (I-82).
/// </summary>
public interface IStreamingCell
{
    /// <summary>The cell's canonical <c>A1</c> address.</summary>
    string Address { get; }
    /// <summary>The cell's 1-based row index.</summary>
    int RowIndex { get; }
    /// <summary>The cell's 1-based column index.</summary>
    int ColumnIndex { get; }
    /// <summary>The cell's kind.</summary>
    CellKind Kind { get; }

    /// <summary>
    /// Writes a string value. Streaming cells store plain strings only —
    /// the rich-text surface from <see cref="ICell.SetRichText"/> is
    /// deliberately absent here. The restriction originated with the v1
    /// streaming engine (NPOI's SXSSF dropped per-run fonts at flush
    /// time) and is retained at v2.0.0; offering it would be an
    /// additive surface decision. Decision I-50 / decision #7
    /// (streaming type-honesty).
    /// </summary>
    void SetString(string value);
    /// <summary>Writes a double value.</summary>
    void SetNumber(double value);
    /// <summary>Writes a decimal value.</summary>
    void SetNumber(decimal value);
    /// <summary>Writes an int value.</summary>
    void SetNumber(int value);
    /// <summary>Writes a long value.</summary>
    void SetNumber(long value);
    /// <summary>Writes a bool value.</summary>
    void SetBool(bool value);
    /// <summary>Writes a <see cref="DateTime"/> value.</summary>
    void SetDate(DateTime value);
    /// <summary>Writes a formula. Cached value is not pre-computed (design §7.8).</summary>
    void SetFormula(string formula);

    /// <summary>Applies a style (merge semantics; pool-deduped).</summary>
    IStreamingCell Style(CellStyle style);

    /// <summary>Sets the number-format string only.</summary>
    IStreamingCell NumberFormat(string format);
}
