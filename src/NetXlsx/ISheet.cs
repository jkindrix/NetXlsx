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
/// Excel error code, surfaced via <see cref="ICell.GetError"/>
/// (decision #49 / design §6.4). Matches the eight standard Excel
/// formula errors.
/// </summary>
public enum CellError
{
    /// <summary><c>#NULL!</c> — intersection of two ranges that do not intersect.</summary>
    Null,
    /// <summary><c>#DIV/0!</c> — division by zero.</summary>
    DivByZero,
    /// <summary><c>#VALUE!</c> — wrong type of operand.</summary>
    Value,
    /// <summary><c>#REF!</c> — reference is not valid.</summary>
    Ref,
    /// <summary><c>#NAME?</c> — unrecognized name in a formula.</summary>
    Name,
    /// <summary><c>#NUM!</c> — invalid numeric value.</summary>
    Num,
    /// <summary><c>#N/A</c> — value is not available.</summary>
    NotAvailable,
    /// <summary><c>#GETTING_DATA</c> — data fetching error from external source.</summary>
    GettingData,
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
    /// Adds a named range to the workbook. <paramref name="name"/> must be
    /// a valid Excel name (letters / digits / underscores / periods;
    /// must start with a letter or underscore; cannot collide with an
    /// existing cell reference like <c>A1</c>). <paramref name="formula"/>
    /// is the body of the reference (e.g. <c>"Sheet1!$A$1:$B$10"</c>) —
    /// no leading <c>=</c>. <paramref name="sheetScope"/> selects the
    /// sheet the name is scoped to; <c>null</c> (the default, decision I9)
    /// scopes it workbook-wide.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="name"/> or <paramref name="formula"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="name"/> violates Excel's name rules or already exists at the requested scope.</exception>
    /// <exception cref="SheetNameException"><paramref name="sheetScope"/> is non-null and does not match an existing sheet.</exception>
    INamedRange AddNamedRange(string name, string formula, string? sheetScope = null);

    /// <summary>The named ranges currently defined on the workbook (scope-agnostic).</summary>
    System.Collections.Generic.IReadOnlyList<INamedRange> NamedRanges { get; }

    /// <summary>
    /// Protects this workbook against UI-level structural changes
    /// (decision I-54). When <paramref name="options"/> is null,
    /// defaults to <see cref="WorkbookProtection.LockStructure"/>
    /// (the common use case). NPOI 2.7.3 does not expose workbook
    /// password support directly; for password protection, reach
    /// through <see cref="Underlying"/>.
    /// </summary>
    void Protect(WorkbookProtection? options = null);

    /// <summary>
    /// Removes workbook-level protection. No-op if not protected.
    /// </summary>
    void Unprotect();

    /// <summary>
    /// Whether the workbook currently has any structure / windows /
    /// revision lock applied.
    /// </summary>
    bool IsProtected { get; }

    /// <summary>
    /// Escape hatch — direct access to the underlying NPOI <c>XSSFWorkbook</c>
    /// per decision #32. Direct mutation is supported but is not synchronized
    /// with wrapper state; callers using this hatch own the consequences.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFWorkbook Underlying { get; }
}

/// <summary>
/// A named range (a workbook- or sheet-scoped alias for a cell or
/// region). Named ranges are useful for human-readable formula
/// references (<c>=SUM(Sales)</c>) and for documenting intent.
/// </summary>
public interface INamedRange
{
    /// <summary>The range's name as it appears in Excel.</summary>
    string Name { get; }

    /// <summary>
    /// The range body (e.g. <c>"Sheet1!$A$1:$B$10"</c>). No leading
    /// <c>=</c>.
    /// </summary>
    string Formula { get; }

    /// <summary>
    /// The sheet this name is scoped to, or <c>null</c> for a
    /// workbook-scoped name (decision I9).
    /// </summary>
    string? SheetScope { get; }
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
    /// Returns a rectangular range of cells parsed from an A1-style
    /// reference (e.g. <c>"A1:C3"</c>, <c>"A:A"</c>, <c>"1:5"</c>).
    /// Whole-row and whole-column shorthand auto-expand per design §6.10.
    /// </summary>
    /// <exception cref="InvalidCellAddressException">Not a valid range.</exception>
    IRange Range(string a1Range);

    /// <summary>
    /// Returns a rectangular range from explicit 1-based coordinates.
    /// Coordinates are normalized: passing <c>(3, 3, 1, 1)</c> yields
    /// the same range as <c>(1, 1, 3, 3)</c>.
    /// </summary>
    /// <exception cref="System.ArgumentOutOfRangeException">Any coordinate out of Excel's grid.</exception>
    IRange Range(int row1, int col1, int row2, int col2);

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
    /// Returns the column at the 1-based <paramref name="index"/>
    /// (<c>1 == "A"</c>). Columns are lazy handles — accessing one
    /// neither materializes cells nor mutates the file.
    /// </summary>
    /// <exception cref="System.ArgumentOutOfRangeException">Index outside <c>[1, CellAddress.MaxColumn]</c>.</exception>
    IColumn Column(int index);

    /// <summary>
    /// Returns the column at the given letter reference (<c>"A"</c>,
    /// <c>"AA"</c>, <c>"XFD"</c>). Case-insensitive; a leading <c>$</c>
    /// is accepted and ignored.
    /// </summary>
    /// <exception cref="InvalidCellAddressException">Not a valid column letter.</exception>
    IColumn Column(string letter);

    /// <summary>
    /// Freezes the top <paramref name="rows"/> rows. Pass 0 to clear an
    /// existing row freeze. Equivalent to <c>FreezePane(rows, 0)</c>.
    /// </summary>
    void FreezeRows(int rows);

    /// <summary>
    /// Freezes the leftmost <paramref name="cols"/> columns. Pass 0 to
    /// clear an existing column freeze. Equivalent to
    /// <c>FreezePane(0, cols)</c>.
    /// </summary>
    void FreezeColumns(int cols);

    /// <summary>
    /// Freezes a top-left pane of <paramref name="rows"/> rows and
    /// <paramref name="cols"/> columns. Replaces any prior freeze on
    /// this sheet. Pass <c>(0, 0)</c> to clear.
    /// </summary>
    void FreezePane(int rows, int cols);

    /// <summary>
    /// Merges the cells in <paramref name="a1Range"/> (e.g. <c>"A1:C3"</c>).
    /// Pre-existing values in non-anchor cells are preserved in the file
    /// (OOXML semantics) but only the anchor (top-left) cell's value is
    /// displayed in Excel.
    /// </summary>
    /// <exception cref="InvalidCellAddressException">Not a valid range.</exception>
    /// <exception cref="System.InvalidOperationException">The range overlaps an existing merged region.</exception>
    void MergeCells(string a1Range);

    /// <summary>
    /// Unmerges the cells in <paramref name="a1Range"/>. No-op if the
    /// exact range is not currently merged (decision §6.4).
    /// </summary>
    void UnmergeCells(string a1Range);

    /// <summary>Currently-merged regions on this sheet, as canonical <c>A1:C3</c> strings.</summary>
    System.Collections.Generic.IReadOnlyList<string> MergedRanges { get; }

    /// <summary>
    /// Whether this sheet is hidden from Excel's tab bar. Mirrors NPOI's
    /// <c>SheetVisibility.Hidden</c>. (NPOI's <c>VeryHidden</c> state —
    /// hidden from VBA — is not modeled here in v1; reach through
    /// <see cref="Underlying"/> if you need it.)
    /// </summary>
    bool Hidden { get; set; }

    /// <summary>Whether the sheet shows gridlines when viewed.</summary>
    bool ShowGridlines { get; set; }

    /// <summary>
    /// Adds an Excel table (<c>ListObject</c>) over <paramref name="a1Range"/>
    /// with header row, codename <paramref name="name"/>, and optional
    /// <paramref name="style"/> (decision I-51).
    /// <para>
    /// The first row of the range must already contain non-empty string
    /// cells — those become the table's column names. <paramref name="name"/>
    /// must follow Excel's name rules (letters / digits / underscores; must
    /// start with a letter or underscore) and be unique workbook-wide.
    /// </para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="a1Range"/> or <paramref name="name"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// The name violates Excel's rules, the name is already taken, the
    /// range is malformed, the header row is missing, or any header cell
    /// is blank or non-string.
    /// </exception>
    ITable AddTable(string a1Range, string name, string? style = null);

    /// <summary>Snapshot of the tables currently defined on this sheet.</summary>
    System.Collections.Generic.IReadOnlyList<ITable> Tables { get; }

    /// <summary>
    /// Non-throwing table lookup by codename (case-insensitive).
    /// </summary>
    bool TryGetTable(string name, [MaybeNullWhen(false)] out ITable table);



    /// <summary>
    /// Applies an AutoFilter (dropdown filter on the header row) over
    /// <paramref name="a1Range"/> (decision I-56). The range must include
    /// the header row as its first row; Excel's filter dropdown appears
    /// on each header cell. Replaces any existing AutoFilter on this sheet.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="a1Range"/> is null.</exception>
    /// <exception cref="InvalidCellAddressException"><paramref name="a1Range"/> is not a valid A1 range.</exception>
    void SetAutoFilter(string a1Range);

    /// <summary>Removes the AutoFilter from this sheet. No-op if none is set.</summary>
    void ClearAutoFilter();

    /// <summary>Whether this sheet currently has an AutoFilter applied.</summary>
    bool HasAutoFilter { get; }

    /// <summary>
    /// The current AutoFilter range (e.g. <c>"A1:D10"</c>), or <c>null</c>
    /// when no AutoFilter is set.
    /// </summary>
    string? AutoFilterRange { get; }

    /// <summary>
    /// Applies <paramref name="validation"/> to the cells in
    /// <paramref name="a1Range"/> (decision I-55). Excel rejects user
    /// input that fails the rule with a default error dialog.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="a1Range"/> or <paramref name="validation"/> is null.</exception>
    /// <exception cref="InvalidCellAddressException"><paramref name="a1Range"/> is not a valid A1 range.</exception>
    void AddValidation(string a1Range, DataValidation validation);

    /// <summary>
    /// Embeds <paramref name="data"/> as an image anchored at
    /// <paramref name="a1Cell"/> (its top-left corner). The image is
    /// rendered at its natural pixel size; resize through
    /// <see cref="IPicture.Underlying"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="a1Cell"/> or <paramref name="data"/> is null.</exception>
    /// <exception cref="InvalidCellAddressException"><paramref name="a1Cell"/> is not a valid A1 reference.</exception>
    IPicture AddPicture(string a1Cell, byte[] data, ImageFormat format);

    /// <summary>
    /// Embeds <paramref name="data"/> as an image, auto-detecting the
    /// format from the leading magic bytes (PNG or JPEG; anything else
    /// throws <see cref="UnsupportedImageFormatException"/>). Anchored
    /// as in <see cref="AddPicture(string, byte[], ImageFormat)"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="a1Cell"/> or <paramref name="data"/> is null.</exception>
    /// <exception cref="InvalidCellAddressException"><paramref name="a1Cell"/> is not a valid A1 reference.</exception>
    /// <exception cref="UnsupportedImageFormatException">The byte buffer is not a recognized PNG or JPEG.</exception>
    IPicture AddPicture(string a1Cell, byte[] data);

    /// <summary>
    /// Protects this sheet against UI editing (decision I-53). When
    /// <paramref name="password"/> is non-null, Excel requires it to
    /// unprotect — but the password is hashed with a weak algorithm,
    /// so this is a UX guard, not real security. When
    /// <paramref name="options"/> is null, defaults to
    /// <see cref="SheetProtection.Default"/> (every action permitted
    /// for non-locked cells).
    /// </summary>
    void Protect(string? password = null, SheetProtection? options = null);

    /// <summary>
    /// Removes protection from this sheet (no-op if not protected).
    /// </summary>
    void Unprotect();

    /// <summary>
    /// Whether this sheet currently has protection applied.
    /// </summary>
    bool IsProtected { get; }

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

    /// <summary>Whether this row is hidden in Excel.</summary>
    bool Hidden { get; set; }

    /// <summary>
    /// Escape hatch — direct access to the underlying NPOI <c>XSSFRow</c>.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFRow Underlying { get; }
}

/// <summary>
/// A rectangular range of cells on an <see cref="ISheet"/>.
/// Enumeration is sparse by default (only currently-populated cells);
/// use <see cref="EnumerateAll"/> for dense iteration.
/// </summary>
public interface IRange : System.Collections.Generic.IEnumerable<ICell>
{
    /// <summary>
    /// Canonical A1 form of the range — <c>A1:C3</c> for bounded
    /// ranges, single-cell form for 1×1 ranges. Whole-row and
    /// whole-column shorthand expands to the canonical bounded form
    /// per design §6.10.
    /// </summary>
    string Address { get; }

    /// <summary>1-based top row.</summary>
    int FirstRow { get; }
    /// <summary>1-based bottom row (inclusive).</summary>
    int LastRow { get; }
    /// <summary>1-based leftmost column.</summary>
    int FirstCol { get; }
    /// <summary>1-based rightmost column (inclusive).</summary>
    int LastCol { get; }

    /// <summary>The total number of cell coordinates in the rectangle (dense count).</summary>
    int Count { get; }

    /// <summary>The owning sheet.</summary>
    ISheet Sheet { get; }

    /// <summary>
    /// Yields every cell coordinate in the rectangle, including blanks.
    /// Lazily materializes empty cells on demand. For whole-row /
    /// whole-column ranges this can be 1M+ items — sparse base
    /// enumeration via <c>foreach</c> is the usual idiom.
    /// </summary>
    System.Collections.Generic.IEnumerable<ICell> EnumerateAll();

    /// <summary>
    /// Sets every cell in the rectangle to <paramref name="value"/>.
    /// Dispatched on runtime type: string / bool / numeric (int, long,
    /// float, double, decimal) / DateTime / DateOnly / TimeOnly /
    /// TimeSpan. Null clears the cell. Unsupported types throw
    /// <see cref="System.ArgumentException"/>.
    /// </summary>
    IRange Value(object? value);

    /// <summary>Applies <paramref name="style"/> to every cell in the rectangle (dense).</summary>
    IRange Apply(CellStyle style);

    /// <summary>
    /// Merges this range. Shorthand for
    /// <c>sheet.MergeCells(range.Address)</c>; same semantics
    /// (decision §6.4): 1×1 is a no-op; overlap with an existing merge
    /// throws.
    /// </summary>
    IRange Merge();

    /// <summary>
    /// Clears the value of every cell in the rectangle. Styles are
    /// preserved. Inverse of <see cref="Value"/> with a non-null arg.
    /// </summary>
    IRange ClearContents();
}

/// <summary>
/// A single column on an <see cref="ISheet"/>. Columns are lightweight
/// handles — accessing a column does not materialize cells or mutate the
/// underlying file.
/// </summary>
public interface IColumn
{
    /// <summary>The 1-based column index (<c>"A" == 1</c>).</summary>
    int Index { get; }

    /// <summary>The canonical column letter (<c>1 → "A"</c>, <c>27 → "AA"</c>).</summary>
    string Letter { get; }

    /// <summary>The owning sheet.</summary>
    ISheet Sheet { get; }

    /// <summary>
    /// Whether this column is hidden in Excel. Setter takes effect
    /// immediately; reading reflects the current NPOI state.
    /// </summary>
    bool Hidden { get; set; }

    /// <summary>
    /// Column width in Excel "character" units. Setting writes through
    /// to NPOI's 256ths-of-a-character integer representation.
    /// </summary>
    double WidthUnits { get; set; }

    /// <summary>Fluent form of the <see cref="WidthUnits"/> setter.</summary>
    /// <exception cref="System.ArgumentOutOfRangeException"><paramref name="units"/> is negative or NaN.</exception>
    IColumn Width(double units);

    /// <summary>
    /// Sizes this column to fit its populated contents (delegates to
    /// NPOI's <c>AutoSizeColumn</c>). On headless Linux systems
    /// without an OS font stack (libgdiplus / fallback fonts) this
    /// throws <see cref="MissingFontException"/> with installation
    /// guidance (design decision I3).
    /// </summary>
    /// <exception cref="MissingFontException">Font metrics unavailable in this environment.</exception>
    IColumn AutoSize();

    /// <summary>
    /// Applies <paramref name="apply"/> to every populated cell in this
    /// column, top to bottom. Empty cells are skipped (sparse iteration).
    /// </summary>
    IColumn ForEachPopulated(Action<ICell> apply);

    /// <summary>
    /// Sets <paramref name="style"/> as the column's default style. New
    /// cells in this column inherit it; existing cells are unaffected
    /// (NPOI's <c>SetDefaultColumnStyle</c> semantics).
    /// </summary>
    IColumn SetDefaultStyle(CellStyle style);
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

    /// <summary>
    /// Writes a multi-run rich-text value to the cell (decision I-50).
    /// The cell's <see cref="Kind"/> becomes <see cref="CellKind.String"/>;
    /// <see cref="GetString"/> returns the concatenated plain text.
    /// <see cref="GetRichText"/> returns the supplied value back.
    /// Per-run <see cref="RichTextStyle"/> axes that are null inherit the
    /// cell's current font; cell-level fills, borders, and alignment
    /// remain governed by <see cref="Style(CellStyle)"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is null.</exception>
    /// <exception cref="ResourceLimitExceededException">
    /// The concatenated plain text exceeds <c>WorkbookOptions.MaxCellTextLength</c>.
    /// </exception>
    void SetRichText(RichText value);

    /// <summary>
    /// Returns the cell's rich-text value when it carries explicit
    /// per-run formatting, or <c>null</c> when the cell holds a plain
    /// string (or any non-string value). A cell set via
    /// <see cref="SetString"/> always returns <c>null</c> here even
    /// though OOXML stores all string cells as rich-text internally.
    /// </summary>
    RichText? GetRichText();

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

    /// <summary>
    /// Stores <paramref name="formula"/> as a formula. A leading <c>=</c>
    /// is optional — both <c>"=A1+B1"</c> and <c>"A1+B1"</c> are
    /// accepted. The cached value is <b>not</b> pre-computed (design
    /// decision #46 / §7.8): Excel and other competent consumers
    /// recalculate on open. Cell kind becomes
    /// <see cref="CellKind.Formula"/>.
    /// </summary>
    /// <exception cref="FormulaException">
    /// The formula body is empty or NPOI cannot parse it.
    /// </exception>
    void SetFormula(string formula);

    /// <summary>
    /// Returns the cell's formula body prefixed with <c>=</c>, or
    /// <c>null</c> when the cell does not hold a formula.
    /// </summary>
    string? GetFormula();

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
    /// Applies <paramref name="style"/> to the cell as a <b>merge</b>, not
    /// a replace: only properties whose value is non-null on
    /// <paramref name="style"/> override the corresponding axis on the
    /// cell's existing style; null axes are left untouched. Concretely,
    /// after <c>cell.SetDate(d); cell.Style(new CellStyle { Bold = true });</c>
    /// the cell carries both <c>Bold = true</c> and the date number
    /// format.
    /// <para>
    /// Because the merge is "existing ← overlay-non-null",
    /// passing <see cref="CellStyle.Default"/> (every axis null) is a
    /// no-op — it does <b>not</b> clear or reset the cell's style. v1
    /// has no explicit clear sentinel; the deterministic alternative is
    /// to assign a brand-new cell value (e.g. <see cref="Clear"/>
    /// followed by a setter) and then style from scratch.
    /// </para>
    /// <para>
    /// Resolved through the workbook's style-pool dedup (decision #4) —
    /// equal merged styles share one underlying NPOI
    /// <c>ICellStyle</c> index.
    /// </para>
    /// </summary>
    ICell Style(CellStyle style);

    /// <summary>
    /// Shortcut for the common case: applies the given Excel number
    /// format string to the cell while leaving other style properties
    /// untouched. Pass-through bytes per §7.2.
    /// </summary>
    ICell NumberFormat(string format);

    /// <summary>Returns the cell's current style as a value record.</summary>
    CellStyle GetStyle();

    /// <summary>
    /// Attaches a comment to the cell. <paramref name="text"/> is the
    /// comment body; <paramref name="author"/> defaults to
    /// <c>"NetXlsx"</c> per decision I11 (avoids leaking
    /// <c>Environment.UserName</c>). Replaces any existing comment on
    /// the cell. Returns the cell for chaining.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is null.</exception>
    ICell Comment(string text, string? author = null);

    /// <summary>The cell's comment text, or <c>null</c> when the cell carries no comment.</summary>
    string? GetComment();

    /// <summary>The cell's comment author, or <c>null</c> when the cell carries no comment.</summary>
    string? GetCommentAuthor();

    /// <summary>
    /// Attaches a hyperlink to the cell. <paramref name="target"/> is
    /// scheme-sniffed (decision I13): supported schemes are
    /// <c>http://</c>, <c>https://</c>, <c>mailto:</c>, <c>file://</c>,
    /// and the internal <c>#Sheet!Range</c> form. Anything else throws
    /// <see cref="ArgumentException"/>.
    /// <para>
    /// If <paramref name="display"/> is supplied, the cell's displayed
    /// string is set to it (replacing any existing value). If
    /// <paramref name="display"/> is null and the cell is currently
    /// empty, the cell's displayed value is set to
    /// <paramref name="target"/>.
    /// </para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="target"/> is null.</exception>
    /// <exception cref="ArgumentException">Target uses an unsupported scheme.</exception>
    ICell Hyperlink(string target, string? display = null);

    /// <summary>The cell's hyperlink target, or <c>null</c> when no hyperlink is attached.</summary>
    string? GetHyperlink();

    /// <summary>
    /// Returns the cell's Excel error code if it is an error cell (or a
    /// formula cell whose cached result is an error), otherwise <c>null</c>.
    /// </summary>
    CellError? GetError();

    /// <summary>
    /// Escape hatch — direct access to the underlying NPOI <c>XSSFCell</c>.
    /// See <see cref="IWorkbook.Underlying"/> for the contract.
    /// </summary>
    NPOI.XSSF.UserModel.XSSFCell Underlying { get; }
}
