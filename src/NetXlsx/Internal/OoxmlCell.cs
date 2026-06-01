// I-82 engine swap — Open XML SDK-backed ICell (cells & rows slice).
//
// Lazy materialization model (decision #40): an OoxmlCell is a (sheet, row,
// col) handle. Reads look up the existing <c> element and report
// CellKind.Empty when it is absent — accessing a never-written address does
// NOT add a node to the DOM. Writes materialize the <row>/<c> elements (in the
// ascending order Excel requires) via the sheet's grid helpers.
//
// Implemented this slice: string / number / bool set + get, Kind, Clear, the
// numeric-interpretation getters (GetTime/GetDuration), and the address members.
// Deferred (throw NotYet): SetDate/SetTime/SetDuration and SetFormula — a date
// in OOXML is a number plus a date number-format style, so it belongs with the
// styles slice; formulas belong with their own slice. Styling, comments, and
// hyperlinks land in later slices. GetDate/GetFormula/GetError/GetRichText
// return null here, which is correct: this engine cannot yet produce date,
// formula, error, or rich-text cells.

using System;
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
            return KindOf(Element);
        }
    }

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

    // GetDate is null on this engine: a date requires a date number-format
    // style, which this slice does not write (styles slice). Correct, not a stub.
    public DateTime? GetDate() { Wb.ThrowIfDisposed(); return null; }
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
    // A plain string cell returns null per the GetRichText contract.
    public RichText? GetRichText() { Wb.ThrowIfDisposed(); return null; }

    private string ReadString(S.Cell c)
    {
        var t = c.DataType?.Value;
        if (t == S.CellValues.SharedString)
        {
            if (int.TryParse(c.CellValue?.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idx))
                return ResolveSharedString(idx) ?? string.Empty;
            return string.Empty;
        }
        if (t == S.CellValues.InlineString)
            return c.InlineString?.InnerText ?? string.Empty;
        // t == String (formula-string cached) or anything else: raw value text.
        return c.CellValue?.InnerText ?? string.Empty;
    }

    private string? ResolveSharedString(int index)
    {
        var sst = Wb.OpenXmlDocument?.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
        return sst?.Elements<S.SharedStringItem>().ElementAtOrDefault(index)?.InnerText;
    }

    private static double ReadNumber(S.Cell c)
    {
        var text = c.CellValue?.InnerText;
        return double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double v)
            ? v
            : 0.0;
    }

    private static bool ReadBool(S.Cell c)
    {
        var v = c.CellValue?.InnerText;
        return v == "1" || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase);
    }

    // ---- Deferred surface (lands slice by slice; see I-82) -----------------

    private static NotImplementedException NotYet([CallerMemberName] string? member = null)
        => new(
            $"ICell.{member} is not yet implemented on the Open XML SDK engine " +
            "(I-82 engine swap). It lands in a later slice (dates/formulas/styles/" +
            "comments/hyperlinks/rich text); track the swap in docs/design.md (I-82).");

    public void SetRichText(RichText value) => throw NotYet();
    public void SetDate(DateTime value) => throw NotYet();
    public void SetDate(DateOnly value) => throw NotYet();
    public void SetTime(TimeOnly value) => throw NotYet();
    public void SetDuration(TimeSpan value) => throw NotYet();
    public void SetFormula(string formula) => throw NotYet();

    public ICell ApplyNamedStyle(string name) => throw NotYet();
    public ICell Style(CellStyle style) => throw NotYet();
    public ICell NumberFormat(string format) => throw NotYet();
    public CellStyle GetStyle() => throw NotYet();
    public ICell Comment(string text, string? author = null) => throw NotYet();
    public string? GetComment() => throw NotYet();
    public string? GetCommentAuthor() => throw NotYet();
    public ICell Hyperlink(string target, string? display = null) => throw NotYet();
    public string? GetHyperlink() => throw NotYet();

    public NPOI.XSSF.UserModel.XSSFCell Underlying => throw new NotSupportedException(
        "ICell.Underlying (NPOI XSSFCell) is not available on the Open XML SDK " +
        "engine (I-82). Use IWorkbook.OpenXmlDocument for the SDK escape hatch.");
}
