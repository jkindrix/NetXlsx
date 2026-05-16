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
    /// Escape hatch — direct access to the underlying NPOI <c>XSSFSheet</c>.
    /// See <see cref="IWorkbook.Underlying"/> for the contract.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFSheet Underlying { get; }
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

    /// <summary>Writes a boolean value to the cell.</summary>
    void SetBool(bool value);

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
    /// Escape hatch — direct access to the underlying NPOI <c>XSSFCell</c>.
    /// See <see cref="IWorkbook.Underlying"/> for the contract.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFCell Underlying { get; }
}
