// Workbook / Sheet / Cell interfaces per design §6.2-§6.5.
// v0.2.0 vertical-slice subset: enough surface for
//   Workbook.Create -> AddSheet -> sheet["A1"].SetString -> SaveAsync.
// Additional members (rows, ranges, freeze panes, streaming, etc.) land
// in subsequent v0.x milestones.

using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace NetXlsx;

/// <summary>
/// Classification of a cell's stored value. Mirrors NPOI's
/// <c>CellType</c> with the addition of <c>Date</c>, which OOXML stores
/// as a numeric value with a date-style number format.
/// </summary>
public enum CellKind
{
    /// <summary>Cell has no value (blank).</summary>
    Empty,
    /// <summary>Cell holds a string.</summary>
    String,
    /// <summary>Cell holds an IEEE-754 double, with no date number format applied.</summary>
    Number,
    /// <summary>Cell holds a numeric value styled with a date number format.</summary>
    Date,
    /// <summary>Cell holds a boolean.</summary>
    Bool,
    /// <summary>Cell holds a formula; the cached result follows the kind of its computed value.</summary>
    Formula,
    /// <summary>Cell holds an Excel error code (e.g. <c>#DIV/0!</c>, <c>#REF!</c>).</summary>
    Error,
}

/// <summary>
/// Represents an Excel workbook.
/// </summary>
public interface IWorkbook : IDisposable
{
    /// <summary>Number of sheets in the workbook.</summary>
    int SheetCount { get; }

    /// <summary>Looks up a sheet by name. Throws if not found.</summary>
    /// <param name="name">Sheet name (case-insensitive, decision #14 / I5 culture rule).</param>
    ISheet this[string name] { get; }

    /// <summary>Looks up a sheet by 0-based index.</summary>
    ISheet this[int index] { get; }

    /// <summary>
    /// Adds a new sheet to the end of the workbook. The supplied name
    /// must be 1..31 characters, must not contain <c>\ / ? * [ ]</c>, and
    /// must be unique within the workbook (case-insensitive). Throws
    /// <see cref="SheetNameException"/> on rule violation.
    /// </summary>
    ISheet AddSheet(string name);

    /// <summary>Non-throwing sheet lookup.</summary>
    bool TryGetSheet(string name, [MaybeNullWhen(false)] out ISheet sheet);

    /// <summary>Saves the workbook synchronously to <paramref name="stream"/>.</summary>
    void Save(Stream stream, bool leaveOpen = true);

    /// <summary>Saves the workbook synchronously to a file path.</summary>
    void Save(string path);

    /// <summary>Saves the workbook asynchronously to <paramref name="stream"/>.</summary>
    Task SaveAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default);

    /// <summary>Saves the workbook asynchronously to a file path.</summary>
    Task SaveAsync(string path, CancellationToken ct = default);

    /// <summary>
    /// Escape hatch — direct access to the underlying NPOI <c>XSSFWorkbook</c>
    /// per decision #32. Direct mutation is supported but is not synchronized
    /// with wrapper state; callers using this hatch own the consequences.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFWorkbook Underlying { get; }
}

/// <summary>
/// Represents a worksheet within an <see cref="IWorkbook"/>.
/// </summary>
public interface ISheet
{
    /// <summary>The sheet's name.</summary>
    string Name { get; }

    /// <summary>The owning workbook.</summary>
    IWorkbook Workbook { get; }

    /// <summary>
    /// Looks up a cell by <c>A1</c> address. Returns a materialized
    /// <see cref="ICell"/> even for never-written addresses (decision #40);
    /// the returned cell reports <see cref="CellKind.Empty"/>.
    /// </summary>
    /// <exception cref="InvalidCellAddressException">The address is not a valid <c>A1</c> reference.</exception>
    ICell this[string a1] { get; }

    /// <summary>
    /// Looks up a cell by 1-based row and column (decision #3). Equivalent
    /// to <c>sheet[CellAddress.Format(row, column)]</c> but skips parsing.
    /// </summary>
    /// <exception cref="System.ArgumentOutOfRangeException">Row or column out of Excel's grid limits.</exception>
    ICell this[int row, int column] { get; }

    /// <summary>
    /// Appends a new row immediately after the last written row. For an
    /// empty sheet, creates row 1. Idempotent w.r.t. the underlying NPOI
    /// row index (never overwrites an existing row).
    /// </summary>
    IRow AppendRow();

    /// <summary>
    /// Returns the row at the 1-based <paramref name="index"/>, materializing
    /// an empty row if none exists at that index. The returned row's cells
    /// follow decision #40 (auto-materialize as empty on access).
    /// </summary>
    /// <exception cref="System.ArgumentOutOfRangeException">Row index out of Excel's row range.</exception>
    IRow Row(int index);

    /// <summary>
    /// Escape hatch — direct access to the underlying NPOI <c>XSSFSheet</c>.
    /// See <see cref="IWorkbook.Underlying"/> for the contract.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFSheet Underlying { get; }
}

/// <summary>
/// Represents a row within an <see cref="ISheet"/>. Cells are 1-based
/// (decision #3); fluent setters return the row itself for chaining.
/// </summary>
public interface IRow
{
    /// <summary>The row's 1-based index.</summary>
    int Index { get; }

    /// <summary>The owning sheet.</summary>
    ISheet Sheet { get; }

    /// <summary>
    /// Returns the cell at <paramref name="column"/> (1-based), materializing
    /// an empty cell if none exists.
    /// </summary>
    /// <exception cref="System.ArgumentOutOfRangeException">Column out of range.</exception>
    ICell Cell(int column);

    /// <summary>Indexer form of <see cref="Cell(int)"/>.</summary>
    ICell this[int column] { get; }

    /// <summary>Indexer keyed by column letter (e.g. <c>"A"</c>, <c>"AA"</c>).</summary>
    /// <exception cref="System.ArgumentException">Not a valid column letter.</exception>
    ICell this[string columnLetter] { get; }

    /// <summary>Writes a string value to the column and returns this row for chaining.</summary>
    IRow Set(int column, string value);
    /// <summary>Writes a double value to the column and returns this row for chaining.</summary>
    IRow Set(int column, double value);
    /// <summary>Writes a decimal value to the column and returns this row for chaining.</summary>
    IRow Set(int column, decimal value);
    /// <summary>Writes an int value to the column and returns this row for chaining.</summary>
    IRow Set(int column, int value);
    /// <summary>Writes a long value to the column and returns this row for chaining.</summary>
    IRow Set(int column, long value);
    /// <summary>Writes a bool value to the column and returns this row for chaining.</summary>
    IRow Set(int column, bool value);
    /// <summary>Writes a <see cref="DateTime"/> value (decisions I-17, I-18).</summary>
    IRow Set(int column, DateTime value);
    /// <summary>Writes a <see cref="DateOnly"/> value (decision I-19).</summary>
    IRow Set(int column, DateOnly value);
    /// <summary>Writes a <see cref="TimeOnly"/> value as a fraction of a day.</summary>
    IRow Set(int column, TimeOnly value);
    /// <summary>Writes a <see cref="TimeSpan"/> value as elapsed time.</summary>
    /// <exception cref="ArgumentOutOfRangeException">Negative <paramref name="value"/> (decision I15).</exception>
    IRow Set(int column, TimeSpan value);

    /// <summary>
    /// Escape hatch — direct access to the underlying NPOI <c>XSSFRow</c>.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFRow Underlying { get; }
}

/// <summary>
/// Represents a single cell within an <see cref="ISheet"/>.
/// </summary>
public interface ICell
{
    /// <summary>The cell's canonical <c>A1</c> address (uppercase, no <c>$</c>).</summary>
    string Address { get; }

    /// <summary>The cell's 1-based row index.</summary>
    int RowIndex { get; }

    /// <summary>The cell's 1-based column index. <c>A</c> = 1.</summary>
    int ColumnIndex { get; }

    /// <summary>Classification of the cell's stored value.</summary>
    CellKind Kind { get; }

    /// <summary>Writes a string value to the cell.</summary>
    void SetString(string value);

    /// <summary>Writes a numeric value to the cell.</summary>
    void SetNumber(double value);

    /// <summary>
    /// Writes a decimal value to the cell. Stored as IEEE-754 double per
    /// decision I3.6 / §7.4 — precision loss is possible for values with
    /// more than ~15 significant digits.
    /// </summary>
    void SetNumber(decimal value);

    /// <summary>Writes a 32-bit signed integer value to the cell.</summary>
    void SetNumber(int value);

    /// <summary>
    /// Writes a 64-bit signed integer value to the cell. Values exceeding
    /// the <c>double</c> exact-integer range (±2^53) lose precision when
    /// round-tripped, since OOXML stores numbers as IEEE-754.
    /// </summary>
    void SetNumber(long value);

    /// <summary>Writes a boolean value to the cell.</summary>
    void SetBool(bool value);

    /// <summary>
    /// Writes a <see cref="DateTime"/> value to the cell. The cell is given
    /// the workbook's default date-time number format if it carries no
    /// explicit style (decision I-18). <see cref="DateTime.Kind"/> is
    /// stored as-is — no timezone conversion (decision I17).
    /// </summary>
    void SetDate(DateTime value);

    /// <summary>
    /// Writes a <see cref="DateOnly"/> value to the cell. The cell is given
    /// the workbook's default date number format if it carries no explicit
    /// style (decisions I-18, I-19; ISO <c>yyyy-mm-dd</c>).
    /// </summary>
    void SetDate(DateOnly value);

    /// <summary>
    /// Writes a <see cref="TimeOnly"/> value to the cell as a fraction of
    /// a day. The cell is given the workbook's default time format
    /// (<c>h:mm:ss</c>) if it carries no explicit style.
    /// </summary>
    void SetTime(TimeOnly value);

    /// <summary>
    /// Writes a <see cref="TimeSpan"/> value to the cell as elapsed time.
    /// The cell is given the workbook's default elapsed-time format
    /// (<c>[h]:mm:ss</c>, decision §7.9) if it carries no explicit style.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="value"/> is negative — Excel cannot render negative
    /// times (decision I15).
    /// </exception>
    void SetDuration(TimeSpan value);

    /// <summary>Clears the cell's value and resets its kind to <see cref="CellKind.Empty"/>.</summary>
    void Clear();

    /// <summary>
    /// Returns the cell's string representation (per design §7.10).
    /// Returns the stored string for <see cref="CellKind.String"/>,
    /// invariant-formatted number for numeric, etc. Empty cells return
    /// the empty string.
    /// </summary>
    string GetString();

    /// <summary>
    /// Returns the cell's numeric value, or <c>null</c> for non-numeric
    /// cells. Booleans return <c>1.0</c> / <c>0.0</c> per design I-5.
    /// </summary>
    double? GetNumber();

    /// <summary>
    /// Returns the cell's boolean value, or <c>null</c> for non-boolean
    /// cells.
    /// </summary>
    bool? GetBool();

    /// <summary>
    /// Returns the cell's value as <see cref="DateTime"/>, or <c>null</c>
    /// for non-date cells. Result <see cref="DateTime.Kind"/> is always
    /// <see cref="DateTimeKind.Unspecified"/> (decision I17).
    /// </summary>
    DateTime? GetDate();

    /// <summary>
    /// Returns the cell's value as <see cref="DateOnly"/>, or <c>null</c>
    /// for non-date cells (time-of-day component is dropped).
    /// </summary>
    DateOnly? GetDateOnly();

    /// <summary>
    /// Returns the cell's value as <see cref="TimeOnly"/>, or <c>null</c>
    /// for non-numeric cells and for fractional-day values outside the
    /// <c>[0, 1)</c> range that <see cref="TimeOnly"/> covers (§7.9).
    /// </summary>
    TimeOnly? GetTime();

    /// <summary>
    /// Returns the cell's value as <see cref="TimeSpan"/>, or <c>null</c>
    /// for non-numeric cells. Interprets the cell's numeric value as
    /// elapsed days (§7.9).
    /// </summary>
    TimeSpan? GetDuration();

    /// <summary>
    /// Escape hatch — direct access to the underlying NPOI <c>XSSFCell</c>.
    /// See <see cref="IWorkbook.Underlying"/> for the contract.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFCell Underlying { get; }
}
