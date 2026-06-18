// ICell — single-cell surface, plus CellKind / CellError enums that are
// intrinsically cell-level concepts. Split out of ISheet.cs at v1.2 /
// v1.1-review item 2.

using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace NetXlsx;

/// <summary>
/// Classification of a cell's stored value. Mirrors the classic
/// POI/NPOI <c>CellType</c> taxonomy with the addition of <c>Date</c>,
/// which OOXML stores as a numeric value with a date-style number
/// format.
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

    /// <summary>Writes a string value to the cell.
    /// <para>
    /// Characters XML 1.0 cannot carry (0x00–0x08, 0x0B, 0x0C, 0x0E–0x1F,
    /// U+FFFE/U+FFFF, lone surrogate halves) plus CR (0x0D, which XML
    /// writers and readers silently normalize away) are stored using
    /// Excel's own <c>_xHHHH_</c> escape convention (ECMA-376 ST_Xstring,
    /// decision I-88) and decoded transparently by the typed getters —
    /// the round-trip is lossless through NetXlsx itself, Excel, POI,
    /// ClosedXML, exceljs, calamine ≥ 0.31.0, and LibreOffice ≥ 26.x
    /// (oracle-verified). Consumers that do not implement the convention
    /// (openpyxl — and therefore pandas' default engine — and older
    /// calamine) surface the literal escape text such as <c>_x0007_</c>
    /// instead.
    /// </para>
    /// <para>
    /// Editing the cell text does <b>not</b> remove an attached hyperlink:
    /// after <see cref="Hyperlink(string, string?)"/>, a later
    /// <c>SetString</c> changes the displayed text but the link (and its
    /// relationship for an external target) stays bound to the cell — Excel
    /// behaves identically. Call <see cref="RemoveHyperlink"/> to drop it.
    /// </para>
    /// </summary>
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
    /// The formula body is empty or fails structural validation
    /// (unbalanced parentheses, unterminated string literal). Semantic
    /// errors are Excel's to report — no evaluation is attempted.
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
    /// Applies the style previously registered via
    /// <see cref="IWorkbook.RegisterStyle"/> under <paramref name="name"/>
    /// (decision I-57). Equivalent to
    /// <c>cell.Style(workbook.GetRegisteredStyle(name)!)</c> but throws
    /// a friendlier error when the name is unknown.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="name"/> is null.</exception>
    /// <exception cref="ArgumentException">No style is registered under <paramref name="name"/>.</exception>
    ICell ApplyNamedStyle(string name);

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
    /// equal merged styles share one underlying stylesheet
    /// <c>&lt;xf&gt;</c> (cellXfs) entry.
    /// </para>
    /// </summary>
    ICell Style(CellStyle style);

    /// <summary>
    /// Shortcut for the common case: applies the given Excel number
    /// format string to the cell while leaving other style properties
    /// untouched. Pass-through bytes per §7.2.
    /// <para>
    /// The format code is <b>not validated</b> (R-16, deliberate): the
    /// Excel format-string grammar is rendered — and arbitrated — by the
    /// consumer at display time, and a malformed code degrades to General
    /// display rather than corrupting the file. Compile-time
    /// <c>[Column(Format = ...)]</c> codes get a structural smoke check
    /// (<c>NXLS0003</c>); this runtime path is pass-through by design.
    /// </para>
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
    /// Removes the comment attached to the cell (its text, the legacy VML
    /// popup shape, and — when it was the last comment on the sheet — the
    /// comments part itself, leaving no empty artifact). Idempotent: a no-op
    /// returning the cell unchanged when the cell carries no comment.
    /// Returns the cell for chaining.
    /// </summary>
    ICell RemoveComment();

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
    /// <para>
    /// The link binds to the cell, not to its text: a later
    /// <see cref="SetString(string)"/> (or any value edit) changes the
    /// displayed value while the hyperlink — and, for an external target,
    /// its relationship — remains attached, matching Excel. Use
    /// <see cref="RemoveHyperlink"/> to detach it.
    /// </para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="target"/> is null.</exception>
    /// <exception cref="ArgumentException">Target uses an unsupported scheme.</exception>
    ICell Hyperlink(string target, string? display = null);

    /// <summary>The cell's hyperlink target, or <c>null</c> when no hyperlink is attached.</summary>
    string? GetHyperlink();

    /// <summary>
    /// Removes the hyperlink attached to the cell: the <c>&lt;hyperlink&gt;</c>
    /// element and, for an external target, its reference relationship. The
    /// cell's displayed text is left untouched. Idempotent: a no-op returning
    /// the cell unchanged when no hyperlink is attached. Returns the cell for
    /// chaining.
    /// </summary>
    ICell RemoveHyperlink();

    /// <summary>
    /// Returns the cell's Excel error code if it is an error cell (or a
    /// formula cell whose cached result is an error), otherwise <c>null</c>.
    /// </summary>
    CellError? GetError();

    /// <summary>
    /// Escape hatch — direct access to the underlying Open XML SDK
    /// <see cref="DocumentFormat.OpenXml.Spreadsheet.Cell"/> element (I-82).
    /// Reaching for the raw node is a write-like act: on a never-written
    /// address this materializes the cell element (decision #40 lazy cells
    /// stay lazy until the hatch is used). See
    /// <see cref="IWorkbook.Underlying"/> for the contract.
    /// </summary>
    DocumentFormat.OpenXml.Spreadsheet.Cell Underlying { get; }
}
