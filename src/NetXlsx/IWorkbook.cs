// Workbook-level interfaces — IWorkbook (decision §6.2) and INamedRange
// (workbook-scoped named ranges, decision I9). Split out of ISheet.cs
// at v1.2 / v1.1-review item 2.

using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace NetXlsx;

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
    /// Returns a snapshot of the workbook's style-pool dedup activity
    /// (decision I-61). Useful for ops verification — confirms the
    /// pool is sharing styles / fonts rather than allocating per cell.
    /// The returned struct is a snapshot; subsequent mutations don't
    /// update it.
    /// </summary>
    StylePoolDiagnostics GetStylePoolDiagnostics();

    /// <summary>
    /// Registers <paramref name="style"/> under <paramref name="name"/> for
    /// reuse via <see cref="ICell.ApplyNamedStyle"/> and
    /// <see cref="IRange.ApplyNamedStyle"/> (decision I-57). Names are
    /// case-insensitive. Re-registering an existing name replaces the
    /// definition.
    /// <para>
    /// v1.1 named styles are an <b>in-process convenience</b> — they do
    /// not produce entries in OOXML's named-style table. Reading a
    /// saved workbook with <see cref="Workbook.Open"/> does not rehydrate
    /// the name map; only the per-cell style is preserved (via the
    /// style-pool dedup, decision #4).
    /// </para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="name"/> or <paramref name="style"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="name"/> is empty.</exception>
    void RegisterStyle(string name, CellStyle style);

    /// <summary>
    /// Returns the style registered under <paramref name="name"/>, or
    /// <c>null</c> when no such name is registered. Case-insensitive.
    /// </summary>
    CellStyle? GetRegisteredStyle(string name);

    /// <summary>The names currently registered via <see cref="RegisterStyle"/>.</summary>
    System.Collections.Generic.IReadOnlyCollection<string> RegisteredStyleNames { get; }

    /// <summary>
    /// Protects this workbook against UI-level structural changes
    /// (decision I-54). When <paramref name="options"/> is null,
    /// defaults to <see cref="WorkbookProtection.LockStructure"/>.
    /// For password protection, use <see cref="ProtectWithPassword"/>
    /// (decision I-65).
    /// </summary>
    void Protect(WorkbookProtection? options = null);

    /// <summary>
    /// Protects this workbook with a password (decision I-65). Excel
    /// requires <paramref name="password"/> before allowing
    /// unprotection through the UI — but the underlying hash is the
    /// legacy 16-bit XOR-verifier algorithm, widely known to be
    /// brute-forceable. This is a UX guard, not security.
    /// <para>
    /// A separate method (rather than an overload of
    /// <see cref="Protect(WorkbookProtection?)"/>) avoids the
    /// call-site ambiguity that two same-named methods with all-
    /// optional parameters would create. Method naming is also more
    /// self-documenting at the call site.
    /// </para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="password"/> is null. Use <see cref="Protect(WorkbookProtection?)"/> for the no-password case.</exception>
    void ProtectWithPassword(string password, WorkbookProtection? options = null);

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
    /// Injects an OOXML theme part (<c>xl/theme/theme1.xml</c>) into the
    /// workbook (decision I-79). The theme defines color scheme, font
    /// scheme, and effect scheme that Excel uses to resolve theme-based
    /// references in cell styles. Without a theme, Excel falls back to
    /// defaults that may not match Excel-authored files for column widths
    /// (font metric resolution) and theme-indexed colors.
    /// <para>
    /// The byte content should be a complete theme1.xml document. NetXlsx
    /// creates the OPC part and the workbook→theme relationship.
    /// </para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="themeXml"/> is null.</exception>
    void SetThemeXml(byte[] themeXml);

    /// <summary>
    /// Returns the workbook's <c>xl/theme/theme1.xml</c> bytes (decision
    /// I-81), or <c>null</c> if the workbook has no theme part.
    /// Counterpart to <see cref="SetThemeXml(byte[])"/>; together they
    /// support round-tripping a workbook's theme as opaque XML.
    /// </summary>
    byte[]? GetThemeXml();

    /// <summary>
    /// Resolves a theme color reference to an explicit RGB
    /// <see cref="Color"/> against this workbook's theme (decision I-81).
    /// The integer encoding follows OOXML's cell-color theme index
    /// (<c>0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2, 4..9 = accent1..6,
    /// 10 = hlink, 11 = folHlink</c>) — matching <see cref="ThemeColor.Index"/>.
    /// <paramref name="tint"/> is applied with Excel's tint algorithm
    /// (negative darkens, positive lightens). Returns <c>null</c> when the
    /// workbook has no theme or the index isn't defined in the scheme.
    /// </summary>
    Color? ResolveThemeColor(int index, double tint = 0);

    /// <summary>Convenience overload for a <see cref="ThemeColor"/>.</summary>
    Color? ResolveThemeColor(ThemeColor color);

    /// <summary>
    /// Resolves a scheme color reference by its drawing name (e.g.
    /// <c>"dk1"</c>, <c>"accent3"</c>, <c>"tx1"</c>) to an explicit RGB
    /// <see cref="Color"/> (decision I-81). Drawing references use the
    /// <c>tx1</c>/<c>bg1</c>/<c>tx2</c>/<c>bg2</c> aliases for
    /// <c>dk1</c>/<c>lt1</c>/<c>dk2</c>/<c>lt2</c>; both spellings are
    /// accepted. <paramref name="tint"/> is applied with Excel's tint
    /// algorithm. Returns <c>null</c> for an unknown name or an absent
    /// theme.
    /// </summary>
    Color? ResolveThemeColor(string schemeName, double tint = 0);

    /// <summary>
    /// EMU line width for a theme line-style reference
    /// (<c>themeElements/fmtScheme/lnStyleLst/ln[idx]/@w</c>), used by
    /// connectors and shapes carrying a <c>style/lnRef</c> (decision I-81).
    /// <paramref name="oneBasedIdx"/> is 1-based to match the
    /// <c>lnRef/@idx</c> attribute. Returns <c>null</c> if the index is
    /// out of range or the theme has no <c>lnStyleLst</c>.
    /// </summary>
    int? GetThemeLineWidthEmu(int oneBasedIdx);

    /// <summary>
    /// Whether this workbook is macro-enabled (<c>.xlsm</c>) per decision
    /// I-69. A macro-enabled workbook preserves VBA project parts across
    /// open/save; NetXlsx does not read, write, or execute VBA — the
    /// macro content is passthrough only.
    /// </summary>
    bool IsMacroEnabled { get; }

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
