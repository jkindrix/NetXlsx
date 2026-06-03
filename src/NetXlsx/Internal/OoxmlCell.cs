// I-82 engine swap — Open XML SDK-backed ICell (cells & rows slice).
//
// Lazy materialization model (decision #40): an OoxmlCell is a (sheet, row,
// col) handle. Reads look up the existing <c> element and report
// CellKind.Empty when it is absent — accessing a never-written address does
// NOT add a node to the DOM. Writes materialize the <row>/<c> elements (in the
// ascending order Excel requires) via the sheet's grid helpers.
//
// Implemented across slices: string / number / bool set + get, Kind, Clear, the
// numeric-interpretation getters (GetTime/GetDuration), the address members
// (cells & rows); date/time setters + styling (cell styles); rich text
// (SetRichText/GetRichText, this slice — inline <r> runs, with empty-style runs
// inheriting the cell font per lesson #10). Deferred (throw NotYet): SetFormula,
// comments, hyperlinks land in later slices. GetFormula/GetError return null
// here, which is correct: this engine cannot yet produce formula or error cells.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlCell : ICell
{
    private readonly OoxmlSheet _sheet;
    private readonly int _row;
    private readonly int _col;

    internal OoxmlCell(OoxmlSheet sheet, int row, int col)
    {
        _sheet = sheet;
        _row = row;
        _col = col;
    }

    private OoxmlWorkbook Wb => _sheet.WorkbookInternal;

    // The backing <c> element, or null when the address has never been written.
    private S.Cell? Element => _sheet.FindCell(_row, _col);

    public string Address { get { Wb.ThrowIfDisposed(); return CellAddress.Format(_row, _col); } }
    public int RowIndex { get { Wb.ThrowIfDisposed(); return _row; } }
    public int ColumnIndex { get { Wb.ThrowIfDisposed(); return _col; } }

    public CellKind Kind
    {
        get
        {
            Wb.ThrowIfDisposed();
            var c = Element;
            var k = KindOf(c);
            // A numeric cell with a date/time number format is a Date (parity
            // with the NPOI engine's DateUtil.IsCellDateFormatted classification).
            return k == CellKind.Number && IsDateFormatted(c) ? CellKind.Date : k;
        }
    }

    // True when the cell's applied style carries a date/time number format.
    private bool IsDateFormatted(S.Cell? c)
        => c is not null && Wb.StylePool.IsDateFormatted(c.StyleIndex?.Value ?? 0);

    private static CellKind KindOf(S.Cell? c)
    {
        if (c is null) return CellKind.Empty;
        if (c.CellFormula is not null) return CellKind.Formula;
        // Open XML SDK 3.x models CellValues as a struct, not a C# enum, so it
        // can't appear in switch/case labels — compare with ==.
        var t = c.DataType?.Value ?? S.CellValues.Number;
        if (t == S.CellValues.InlineString || t == S.CellValues.SharedString || t == S.CellValues.String)
            return CellKind.String;
        if (t == S.CellValues.Boolean) return CellKind.Bool;
        if (t == S.CellValues.Error) return CellKind.Error;
        // Number is the default type; an empty <c/> with no value node is a
        // materialized-but-blank cell.
        return c.CellValue is null && c.InlineString is null
            ? CellKind.Empty
            : CellKind.Number;
    }

    // ---- Setters (implemented) ---------------------------------------------

    public void SetString(string value)
    {
        Wb.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(value);
        int limit = Wb.Options.MaxCellTextLength;
        if (value.Length > limit)
            throw new ResourceLimitExceededException("cell text length", limit, value.Length);

        var c = _sheet.GetOrCreateCell(_row, _col);
        c.RemoveAllChildren();
        c.CellFormula = null;
        c.DataType = S.CellValues.InlineString;
        // Space="preserve" keeps leading/trailing/standalone whitespace intact.
        c.AppendChild(new S.InlineString(new S.Text(value) { Space = SpaceProcessingModeValues.Preserve }));
    }

    public void SetNumber(double value)
    {
        Wb.ThrowIfDisposed();
        WriteNumber(value);
    }

    public void SetNumber(decimal value)
    {
        Wb.ThrowIfDisposed();
        // Decision I3.6 / §7.4: stored as IEEE-754 double; precision loss possible.
        WriteNumber((double)value);
    }

    public void SetNumber(int value)
    {
        Wb.ThrowIfDisposed();
        WriteNumber(value);
    }

    public void SetNumber(long value)
    {
        Wb.ThrowIfDisposed();
        // Values > 2^53 lose precision when stored as IEEE-754 double.
        WriteNumber(value);
    }

    private void WriteNumber(double value)
    {
        var c = _sheet.GetOrCreateCell(_row, _col);
        c.RemoveAllChildren();
        c.CellFormula = null;
        c.DataType = null; // number is the default type (no t attribute)
        // G17 round-trips a double exactly and matches the NPOI engine's
        // invariant numeric rendering.
        c.CellValue = new S.CellValue(value.ToString("G17", CultureInfo.InvariantCulture));
    }

    public void SetBool(bool value)
    {
        Wb.ThrowIfDisposed();
        var c = _sheet.GetOrCreateCell(_row, _col);
        c.RemoveAllChildren();
        c.CellFormula = null;
        c.DataType = S.CellValues.Boolean;
        c.CellValue = new S.CellValue(value ? "1" : "0");
    }

    public void Clear()
    {
        Wb.ThrowIfDisposed();
        var c = Element;
        if (c is null) return;
        c.RemoveAllChildren();
        c.CellFormula = null;
        c.DataType = null;
    }

    // ---- Getters (implemented) ---------------------------------------------

    public string GetString()
    {
        Wb.ThrowIfDisposed();
        var c = Element;
        switch (KindOf(c))
        {
            case CellKind.Empty: return string.Empty;
            case CellKind.String: return ReadString(c!);
            case CellKind.Bool: return ReadBool(c!) ? "TRUE" : "FALSE";
            case CellKind.Number:
                return ReadNumber(c!).ToString("G17", CultureInfo.InvariantCulture);
            default: return string.Empty;
        }
    }

    public double? GetNumber()
    {
        Wb.ThrowIfDisposed();
        var c = Element;
        return KindOf(c) switch
        {
            CellKind.Number => ReadNumber(c!),
            CellKind.Bool => ReadBool(c!) ? 1.0 : 0.0,
            _ => null,
        };
    }

    public bool? GetBool()
    {
        Wb.ThrowIfDisposed();
        var c = Element;
        return KindOf(c) == CellKind.Bool ? ReadBool(c!) : null;
    }

    public DateTime? GetDate()
    {
        Wb.ThrowIfDisposed();
        var c = Element;
        if (KindOf(c) != CellKind.Number || !IsDateFormatted(c)) return null;
        // Result kind is always Unspecified (decision I17 — no timezone math).
        return DateTime.SpecifyKind(Wb.FromSerial(ReadNumber(c!)), DateTimeKind.Unspecified);
    }

    public DateOnly? GetDateOnly() => GetDate() is { } dt ? DateOnly.FromDateTime(dt) : null;

    public TimeOnly? GetTime()
    {
        // §7.9: any numeric cell; null when outside TimeOnly's [0, 1) range.
        var num = GetNumber();
        if (num is null || num.Value < 0.0 || num.Value >= 1.0) return null;
        return TimeOnly.FromTimeSpan(TimeSpan.FromDays(num.Value));
    }

    public TimeSpan? GetDuration()
    {
        var num = GetNumber();
        return num is null ? null : TimeSpan.FromDays(num.Value);
    }

    // No formula / error cells are producible on this engine yet; null is correct.
    public string? GetFormula() { Wb.ThrowIfDisposed(); return null; }
    public CellError? GetError() { Wb.ThrowIfDisposed(); return null; }

    private string ReadString(S.Cell c)
    {
        var t = c.DataType?.Value;
        if (t == S.CellValues.SharedString)
            return ResolveSharedStringItem(c).InnerText;
        if (t == S.CellValues.InlineString)
            return c.InlineString?.InnerText ?? string.Empty;
        // t == String (formula-string cached) or anything else: raw value text.
        return c.CellValue?.InnerText ?? string.Empty;
    }

    // Resolves a shared-string cell's <v> index to its <si> element, failing loud
    // on a malformed reference (non-integer, negative, or out of range) rather than
    // silently substituting "" (decision I-83 — fail-loud parity with the NPOI
    // engine, which throws here). OOXML defines no default for a corrupt
    // shared-string index, so a silent default would hide data corruption — exactly
    // the silent truncation the library's honesty discipline forbids. Reached only
    // after KindOf classified the cell as a (shared) string, so a failure here is
    // genuine file corruption, never a normal type mismatch.
    private S.SharedStringItem ResolveSharedStringItem(S.Cell c)
    {
        var raw = c.CellValue?.InnerText;
        if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idx) || idx < 0)
            throw new MalformedFileException(
                $"cell {Address}: shared-string index '{raw}' is not a non-negative integer");
        return SharedStringItemAt(idx)
            ?? throw new MalformedFileException(
                $"cell {Address}: shared-string index {idx} is out of range");
    }

    // Reads a numeric cell's <v>, failing loud on an unparseable value rather than
    // silently substituting 0 (decision I-83 — fail-loud parity with the NPOI
    // engine). Only reached once KindOf has classified the cell as Number, so a
    // parse failure here is genuine file corruption, never a normal type mismatch.
    private double ReadNumber(S.Cell c)
    {
        var text = c.CellValue?.InnerText;
        if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double v))
            return v;
        throw new MalformedFileException($"cell {Address}: numeric value '{text}' is not a valid number");
    }

    private static bool ReadBool(S.Cell c)
    {
        var v = c.CellValue?.InnerText;
        return v == "1" || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase);
    }

    // ---- Date / time setters (styles slice) --------------------------------
    // A date in OOXML is a numeric serial plus a date number-format style. Each
    // setter writes the serial and applies the workbook's default date/time/
    // duration format if the cell carries no explicit style (decisions I-18/I-19,
    // §7.9) — mirroring the NPOI engine's XssfCell.Values semantics.

    public void SetDate(DateTime value)
    {
        Wb.ThrowIfDisposed();
        // Decision I17: stored verbatim; no timezone conversion.
        WriteSerial(Wb.ToSerial(value), OoxmlWorkbook.DateTimeStyleSpec);
    }

    public void SetDate(DateOnly value)
    {
        Wb.ThrowIfDisposed();
        WriteSerial(Wb.ToSerial(value.ToDateTime(TimeOnly.MinValue)), OoxmlWorkbook.DateStyleSpec);
    }

    public void SetTime(TimeOnly value)
    {
        Wb.ThrowIfDisposed();
        // Fraction-of-day; no epoch involved (§7.9).
        WriteSerial(value.ToTimeSpan().TotalDays, OoxmlWorkbook.TimeStyleSpec);
    }

    public void SetDuration(TimeSpan value)
    {
        Wb.ThrowIfDisposed();
        if (value < TimeSpan.Zero)
            throw new ArgumentOutOfRangeException(
                nameof(value), value,
                "Negative TimeSpan cannot be stored as Excel duration (decision I15). " +
                "If a signed duration is required, store the value as a number and " +
                "apply a custom format string via the styling API.");
        WriteSerial(value.TotalDays, OoxmlWorkbook.DurationStyleSpec);
    }

    private void WriteSerial(double serial, CellStyle defaultStyle)
    {
        var c = _sheet.GetOrCreateCell(_row, _col);
        c.RemoveAllChildren();
        c.CellFormula = null;
        c.DataType = null; // numeric is the default type
        c.CellValue = new S.CellValue(serial.ToString("G17", CultureInfo.InvariantCulture));
        // Apply the default date/time format only when the cell is unstyled
        // (decision I-18: a user-set style is preserved).
        if ((c.StyleIndex?.Value ?? 0) == 0)
            c.StyleIndex = Wb.StylePool.GetOrCreate(defaultStyle);
    }

    // ---- Rich text (rich-text slice) ---------------------------------------
    // A multi-run formatted string is written as an inline rich string —
    // <c t="inlineStr"><is><r>…</r>…</is></c> — one <r> per RichTextRun. Each
    // run's <rPr> carries its font axes inline (the style pool's run-property
    // builder); a run whose style is empty gets NO <rPr> and so inherits the
    // cell font (lesson #10 — the exact inheritance the NPOI path could not
    // preserve). Kind is String; GetString returns the concatenated run text.

    public void SetRichText(RichText value)
    {
        Wb.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(value);

        var plain = value.PlainText;
        int limit = Wb.Options.MaxCellTextLength;
        if (plain.Length > limit)
            throw new ResourceLimitExceededException("cell text length", limit, plain.Length);

        var c = _sheet.GetOrCreateCell(_row, _col);
        c.RemoveAllChildren();
        c.CellFormula = null;
        c.DataType = S.CellValues.InlineString;

        var inline = new S.InlineString();
        foreach (var run in value.Runs)
        {
            // Empty runs contribute no formatting run (parity with the NPOI engine).
            if (run.Text.Length == 0) continue;
            var r = new S.Run();
            // <rPr> must precede <t> in an <r>; a null <rPr> means "inherit".
            if (OoxmlStylePool.BuildRunProperties(run.Style) is { } rpr) r.AppendChild(rpr);
            r.AppendChild(new S.Text(run.Text) { Space = SpaceProcessingModeValues.Preserve });
            inline.AppendChild(r);
        }
        c.AppendChild(inline);
    }

    public RichText? GetRichText()
    {
        Wb.ThrowIfDisposed();
        var c = Element;
        // Only string cells can carry formatting runs; everything else is null.
        if (KindOf(c) != CellKind.String) return null;

        var runs = ReadRuns(c!);
        // A plain string (no <r> runs) returns null per the GetRichText contract,
        // distinguishing "set via SetRichText / file has rich text" from "plain".
        return runs is null || runs.Count == 0 ? null : new RichText(runs);
    }

    // Collects the <r> formatting runs of a string cell, or null when it carries
    // none (a plain inline/shared <t>). Handles both the inline strings this
    // engine writes and shared-string <si> runs an opened file may carry.
    private List<RichTextRun>? ReadRuns(S.Cell c)
    {
        var t = c.DataType?.Value;
        IEnumerable<S.Run>? rawRuns = null;
        if (t == S.CellValues.InlineString)
        {
            rawRuns = c.InlineString?.Elements<S.Run>();
        }
        else if (t == S.CellValues.SharedString)
        {
            // Fail loud on a corrupt index (decision I-83), consistent with ReadString.
            rawRuns = ResolveSharedStringItem(c).Elements<S.Run>();
        }
        if (rawRuns is null) return null;

        var runs = new List<RichTextRun>();
        foreach (var r in rawRuns)
        {
            var text = r.Text?.InnerText ?? string.Empty;
            runs.Add(new RichTextRun(text, OoxmlStylePool.ReadRunStyle(r.RunProperties)));
        }
        return runs;
    }

    private S.SharedStringItem? SharedStringItemAt(int index)
    {
        var sst = Wb.OpenXmlDocument?.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
        return sst?.Elements<S.SharedStringItem>().ElementAtOrDefault(index);
    }

    // ---- Styling (styles slice) --------------------------------------------

    public ICell Style(CellStyle style)
    {
        Wb.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(style);

        // Merge: non-null axes in `style` overlay the cell's current style; null
        // axes inherit. CellStyle.Default (all-null) is a no-op. Resolved through
        // the pool's dedup so equal merged styles share one cellXfs index (#4).
        var c = _sheet.GetOrCreateCell(_row, _col);
        var current = Wb.StylePool.ReadStyle(c.StyleIndex?.Value ?? 0);
        var merged = Merge(current, style);
        uint idx = Wb.StylePool.GetOrCreate(merged);
        c.StyleIndex = idx == 0 ? null : idx;
        return this;
    }

    public ICell NumberFormat(string format)
    {
        Wb.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(format);
        return Style(new CellStyle { NumberFormat = format });
    }

    public ICell ApplyNamedStyle(string name)
    {
        Wb.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        return Style(Wb.ResolveNamedStyleOrThrow(name));
    }

    public CellStyle GetStyle()
    {
        Wb.ThrowIfDisposed();
        return Wb.StylePool.ReadStyle(Element?.StyleIndex?.Value ?? 0);
    }

    private static CellStyle Merge(CellStyle existing, CellStyle overlay) => new()
    {
        Bold = overlay.Bold ?? existing.Bold,
        Italic = overlay.Italic ?? existing.Italic,
        Underline = overlay.Underline ?? existing.Underline,
        FontName = overlay.FontName ?? existing.FontName,
        FontSize = overlay.FontSize ?? existing.FontSize,
        FontColor = overlay.FontColor ?? existing.FontColor,
        Background = overlay.Background ?? existing.Background,
        BackgroundTheme = overlay.BackgroundTheme ?? existing.BackgroundTheme,
        NumberFormat = overlay.NumberFormat ?? existing.NumberFormat,
        HorizontalAlignment = overlay.HorizontalAlignment ?? existing.HorizontalAlignment,
        VerticalAlignment = overlay.VerticalAlignment ?? existing.VerticalAlignment,
        WrapText = overlay.WrapText ?? existing.WrapText,
        Borders = overlay.Borders ?? existing.Borders,
    };

    // ---- Deferred surface (lands slice by slice; see I-82) -----------------

    private static NotImplementedException NotYet([CallerMemberName] string? member = null)
        => new(
            $"ICell.{member} is not yet implemented on the Open XML SDK engine " +
            "(I-82 engine swap). It lands in a later slice (formulas/comments/" +
            "hyperlinks/rich text); track the swap in docs/design.md (I-82).");

    public void SetFormula(string formula) => throw NotYet();

    public ICell Comment(string text, string? author = null) => throw NotYet();
    public string? GetComment() => throw NotYet();
    public string? GetCommentAuthor() => throw NotYet();
    public ICell Hyperlink(string target, string? display = null) => throw NotYet();
    public string? GetHyperlink() => throw NotYet();

    public NPOI.XSSF.UserModel.XSSFCell Underlying => throw new NotSupportedException(
        "ICell.Underlying (NPOI XSSFCell) is not available on the Open XML SDK " +
        "engine (I-82). Use IWorkbook.OpenXmlDocument for the SDK escape hatch.");
}
