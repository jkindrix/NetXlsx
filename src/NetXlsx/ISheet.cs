// ISheet — the per-worksheet surface (decision §6.4). Split out from the
// monolithic original at v1.2 / v1.1-review item 2 — workbook, row,
// range, column, cell now live in their own files. CellKind + CellError
// enums moved to ICell.cs.

using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace NetXlsx;

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
    /// Groups (outlines) rows <paramref name="startRow"/> through
    /// <paramref name="endRow"/> inclusive (decision I-71). Nested groups
    /// are supported up to Excel's 7-level limit. Grouped rows can be
    /// collapsed in the Excel UI via the outline controls.
    /// </summary>
    /// <param name="startRow">First row of the group (1-based).</param>
    /// <param name="endRow">Last row of the group (1-based, inclusive).</param>
    /// <exception cref="ArgumentOutOfRangeException">Either row index is less than 1 or start &gt; end.</exception>
    void GroupRows(int startRow, int endRow);

    /// <summary>Removes a row group. The range must exactly match a prior <see cref="GroupRows"/> call.</summary>
    void UngroupRows(int startRow, int endRow);

    /// <summary>
    /// Groups (outlines) columns <paramref name="startCol"/> through
    /// <paramref name="endCol"/> inclusive (decision I-71).
    /// </summary>
    /// <param name="startCol">First column of the group (1-based).</param>
    /// <param name="endCol">Last column of the group (1-based, inclusive).</param>
    /// <exception cref="ArgumentOutOfRangeException">Either column index is less than 1 or start &gt; end.</exception>
    void GroupColumns(int startCol, int endCol);

    /// <summary>Removes a column group.</summary>
    void UngroupColumns(int startCol, int endCol);

    /// <summary>
    /// Collapses or expands a row group whose summary row is at
    /// <paramref name="row"/> (1-based). The group must have been
    /// created by a prior <see cref="GroupRows"/> call.
    /// </summary>
    void SetRowGroupCollapsed(int row, bool collapsed);

    /// <summary>
    /// Creates a split (non-frozen) pane on this sheet (decision I-70).
    /// Unlike <see cref="FreezePane"/>, a split pane is draggable by the
    /// user in Excel. Replaces any prior freeze or split on this sheet.
    /// <para>
    /// <paramref name="xSplitTwips"/> is the horizontal split position in
    /// twips (1/20th of a point). <paramref name="ySplitTwips"/> is the
    /// vertical split position in twips. Pass 0 for either dimension to
    /// split in one direction only.
    /// </para>
    /// </summary>
    /// <param name="xSplitTwips">Horizontal split position in twips (0 = no horizontal split).</param>
    /// <param name="ySplitTwips">Vertical split position in twips (0 = no vertical split).</param>
    /// <exception cref="ArgumentOutOfRangeException">Either value is negative.</exception>
    void CreateSplitPane(int xSplitTwips, int ySplitTwips);

    /// <summary>
    /// Adds a shape to this sheet anchored between two cells (decision
    /// I-74). The shape spans from the top-left of <paramref name="startCell"/>
    /// to the bottom-right of <paramref name="endCell"/>.
    /// </summary>
    /// <param name="type">The shape type to draw.</param>
    /// <param name="startCell">Top-left anchor cell in A1 notation.</param>
    /// <param name="endCell">Bottom-right anchor cell in A1 notation.</param>
    /// <param name="fillColor">Optional fill color. Pass <c>null</c> for no fill.</param>
    /// <param name="lineColor">Optional line/border color. Pass <c>null</c> for default black.</param>
    /// <returns>The created shape — use <see cref="IShape.Underlying"/> for advanced properties.</returns>
    IShape AddShape(ShapeType type, string startCell, string endCell, Color? fillColor = null, Color? lineColor = null);

    /// <summary>
    /// Adds one or more conditional formatting rules to the given range
    /// (decision I-73). Each rule is applied in order; Excel evaluates
    /// them top-to-bottom, stopping at the first match per cell.
    /// </summary>
    /// <param name="a1Range">The range to apply formatting to (e.g. <c>"A1:D100"</c>).</param>
    /// <param name="rules">One or more conditional formatting rules.</param>
    /// <exception cref="ArgumentNullException"><paramref name="a1Range"/> or <paramref name="rules"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="rules"/> is empty.</exception>
    void AddConditionalFormatting(string a1Range, params ConditionalFormat[] rules);

    /// <summary>Number of conditional formatting groups on this sheet.</summary>
    int ConditionalFormattingCount { get; }

    /// <summary>Removes the conditional formatting group at <paramref name="index"/> (0-based).</summary>
    void RemoveConditionalFormatting(int index);

    /// <summary>
    /// Sorts rows within <paramref name="a1Range"/> by the specified
    /// <paramref name="keys"/> (decision I-72). The sort is performed
    /// in-memory by physically reordering cell values — no OOXML
    /// <c>sortState</c> metadata is written. Cell styles are moved with
    /// their values.
    /// <para>
    /// The range's first row is <b>not</b> treated as a header; include
    /// only the data rows in the range. Sort keys reference absolute
    /// column indices (1-based), not offsets within the range.
    /// </para>
    /// </summary>
    /// <param name="a1Range">The data range to sort (e.g. <c>"A2:D100"</c>).</param>
    /// <param name="keys">One or more sort keys, applied in order.</param>
    /// <exception cref="ArgumentNullException"><paramref name="a1Range"/> or <paramref name="keys"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="keys"/> is empty.</exception>
    /// <exception cref="InvalidCellAddressException"><paramref name="a1Range"/> is not a valid range.</exception>
    void SortRange(string a1Range, params SortKey[] keys);

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
    /// Removes <paramref name="table"/> from this sheet (decision I-63).
    /// Drops the OOXML <c>&lt;tablePart&gt;</c> entry, the package
    /// relationship, and the underlying table part. The supplied
    /// <see cref="ITable"/> handle becomes stale — subsequent member
    /// access throws <see cref="ObjectDisposedException"/> if the
    /// owning workbook has been disposed, or may produce undefined
    /// values otherwise.
    /// <para>
    /// NPOI 2.7.3's <c>XSSFSheet</c> has no public <c>RemoveTable</c>
    /// method (the source has one upstream but the 2.7.x line never
    /// exposed it). This implementation reaches across the
    /// protected <c>POIXMLDocumentPart.RemoveRelation</c> via
    /// reflection, scoped to that single method. Captured in
    /// <c>docs/implementation-notes.md</c>.
    /// </para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="table"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="table"/> does not belong to this sheet.</exception>
    void RemoveTable(ITable table);



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

    /// <summary>
    /// Applies <paramref name="criteria"/> to the column at
    /// <paramref name="columnOffset"/> within the current AutoFilter
    /// range (decision I-66). <paramref name="columnOffset"/> is
    /// 0-based — column 0 is the first column of the AutoFilter range.
    /// Replaces any existing criteria on that column.
    /// </summary>
    /// <exception cref="InvalidOperationException">No AutoFilter is set on this sheet. Call <see cref="SetAutoFilter"/> first.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnOffset"/> is negative or outside the AutoFilter range.</exception>
    /// <exception cref="ArgumentNullException"><paramref name="criteria"/> is null.</exception>
    void SetAutoFilterColumn(int columnOffset, FilterCriteria criteria);

    /// <summary>
    /// Removes any per-column criteria for the column at
    /// <paramref name="columnOffset"/>. The AutoFilter itself stays in
    /// place (use <see cref="ClearAutoFilter"/> to remove that).
    /// No-op when no criteria is set on the column.
    /// </summary>
    void ClearAutoFilterColumn(int columnOffset);

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
