using System;
using System.Globalization;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfCell : ICell
{
    private readonly XssfWorkbook _workbook;
    private readonly XSSFCell _underlying;
    private readonly int _row1;
    private readonly int _col1;

    public XssfCell(XssfWorkbook workbook, XSSFCell underlying, int row1Based, int col1Based)
    {
        _workbook = workbook;
        _underlying = underlying;
        _row1 = row1Based;
        _col1 = col1Based;
    }

    public string Address
    {
        get { _workbook.ThrowIfDisposed(); return CellAddress.Format(_row1, _col1); }
    }

    public int RowIndex
    {
        get { _workbook.ThrowIfDisposed(); return _row1; }
    }

    public int ColumnIndex
    {
        get { _workbook.ThrowIfDisposed(); return _col1; }
    }

    public CellKind Kind
    {
        get
        {
            _workbook.ThrowIfDisposed();
            return ClassifyKind(_underlying);
        }
    }

    private static CellKind ClassifyKind(XSSFCell cell)
    {
        switch (cell.CellType)
        {
            case CellType.Blank: return CellKind.Empty;
            case CellType.String: return CellKind.String;
            case CellType.Boolean: return CellKind.Bool;
            case CellType.Error: return CellKind.Error;
            case CellType.Formula: return CellKind.Formula;
            case CellType.Numeric:
                return DateUtil.IsCellDateFormatted(cell) ? CellKind.Date : CellKind.Number;
            default: return CellKind.Empty;
        }
    }

    public void SetString(string value)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(value);
        int limit = _workbook.Options.MaxCellTextLength;
        if (value.Length > limit)
        {
            throw new ResourceLimitExceededException(
                "cell text length", limit, value.Length);
        }
        _underlying.SetCellValue(value);
    }

    public void SetNumber(double value)
    {
        _workbook.ThrowIfDisposed();
        _underlying.SetCellValue(value);
    }

    public void SetNumber(decimal value)
    {
        _workbook.ThrowIfDisposed();
        // Decision I3.6 / §7.4: stored as IEEE-754 double; precision loss
        // possible for > ~15 significant digits.
        _underlying.SetCellValue((double)value);
    }

    public void SetNumber(int value)
    {
        _workbook.ThrowIfDisposed();
        _underlying.SetCellValue((double)value);
    }

    public void SetNumber(long value)
    {
        _workbook.ThrowIfDisposed();
        // Values > 2^53 lose precision when stored as IEEE-754 double.
        _underlying.SetCellValue((double)value);
    }

    public void SetBool(bool value)
    {
        _workbook.ThrowIfDisposed();
        _underlying.SetCellValue(value);
    }

    public void SetDate(DateTime value)
    {
        _workbook.ThrowIfDisposed();
        // Decision I17: stored verbatim; no timezone conversion.
        _underlying.SetCellValue(value);
        ApplyDefaultStyleIfUnstyled(_workbook.DateTimeStyle);
    }

    public void SetDate(DateOnly value)
    {
        _workbook.ThrowIfDisposed();
        _underlying.SetCellValue(value.ToDateTime(TimeOnly.MinValue));
        ApplyDefaultStyleIfUnstyled(_workbook.DateStyle);
    }

    public void SetTime(TimeOnly value)
    {
        _workbook.ThrowIfDisposed();
        // TimeOnly is in [00:00:00, 24:00:00) — stored as fraction-of-day.
        var fractionOfDay = value.ToTimeSpan().TotalDays;
        _underlying.SetCellValue(fractionOfDay);
        ApplyDefaultStyleIfUnstyled(_workbook.TimeStyle);
    }

    public void SetDuration(TimeSpan value)
    {
        _workbook.ThrowIfDisposed();
        // Decision I15: Excel cannot render negative time.
        if (value < TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(
                nameof(value), value,
                "Negative TimeSpan cannot be stored as Excel duration (decision I15). " +
                "If a signed duration is required, store the value as a number and " +
                "apply a custom format string via the styling API.");
        }
        _underlying.SetCellValue(value.TotalDays);
        ApplyDefaultStyleIfUnstyled(_workbook.DurationStyle);
    }

    public void SetFormula(string formula)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(formula);

        // Strip an optional leading '=' so callers can write either
        // "=A1+B1" or "A1+B1". NPOI's CellFormula property expects the
        // body without the leading '='.
        var body = formula.Length > 0 && formula[0] == '=' ? formula.Substring(1) : formula;
        if (body.Length == 0)
            throw new FormulaException("formula body is empty (expected '=...' or a non-empty expression)");

        try
        {
            _underlying.SetCellFormula(body);
        }
        catch (Exception ex)
        {
            // NPOI surfaces parse failures as FormulaParseException
            // (NPOI.SS.Formula.FormulaParseException) wrapped in
            // various IllegalArgumentException-style throws. We do not
            // want to leak the NPOI type to callers, so we translate.
            throw new FormulaException(
                $"failed to parse formula '{formula}': {ex.Message}", ex);
        }
        // Per design decision #46 and §7.8: never pre-compute the
        // cached result. Excel recalculates on open.
    }

    public ICell Comment(string text, string? author = null)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(text);

        var factory = _workbook.Underlying.GetCreationHelper();
        var existing = _underlying.CellComment;
        if (existing is not null)
        {
            // NPOI rejects CreateCellComment on a cell that already has
            // one. Mutate in place instead.
            existing.String = factory.CreateRichTextString(text);
            existing.Author = author ?? "NetXlsx";
            return this;
        }

        var sheet = (XSSFSheet)_underlying.Sheet;
        var drawing = sheet.CreateDrawingPatriarch();

        var anchor = factory.CreateClientAnchor();
        anchor.Col1 = _col1 - 1;
        anchor.Row1 = _row1 - 1;
        anchor.Col2 = _col1 + 1;     // 2-column-wide popup
        anchor.Row2 = _row1 + 2;     // 2-row-tall popup

        var comment = drawing.CreateCellComment(anchor);
        comment.String = factory.CreateRichTextString(text);
        comment.Author = author ?? "NetXlsx";   // Decision I11.

        _underlying.CellComment = comment;
        return this;
    }

    public string? GetComment()
    {
        _workbook.ThrowIfDisposed();
        return _underlying.CellComment?.String?.String;
    }

    public string? GetCommentAuthor()
    {
        _workbook.ThrowIfDisposed();
        return _underlying.CellComment?.Author;
    }

    public ICell Hyperlink(string target, string? display = null)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(target);
        if (target.Length == 0)
            throw new ArgumentException("target cannot be empty", nameof(target));

        var (type, address) = SniffHyperlinkScheme(target);

        var link = new NPOI.XSSF.UserModel.XSSFHyperlink(type)
        {
            Address = address,
        };
        if (display is not null) link.Label = display;
        _underlying.Hyperlink = link;

        // If a display value was supplied, set the cell's text to it
        // (replacing any prior value). If not and the cell is empty,
        // fall back to the raw target so the cell isn't blank.
        if (display is not null)
        {
            _underlying.SetCellValue(display);
        }
        else if (_underlying.CellType == NPOI.SS.UserModel.CellType.Blank)
        {
            _underlying.SetCellValue(target);
        }

        return this;
    }

    public string? GetHyperlink()
    {
        _workbook.ThrowIfDisposed();
        return _underlying.Hyperlink?.Address;
    }

    private static (NPOI.SS.UserModel.HyperlinkType type, string address) SniffHyperlinkScheme(string target)
    {
        // Decision I13: scheme-sniff into NPOI's HyperlinkType. We
        // preserve the original target text verbatim — the SchemeType
        // is what tells Excel/OPC how to interpret it.
        if (target.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || target.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            return (NPOI.SS.UserModel.HyperlinkType.Url, target);
        if (target.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
            return (NPOI.SS.UserModel.HyperlinkType.Email, target);
        if (target.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
            return (NPOI.SS.UserModel.HyperlinkType.File, target);
        if (target.StartsWith('#'))
            return (NPOI.SS.UserModel.HyperlinkType.Document, target.Substring(1));
        throw new ArgumentException(
            $"hyperlink target '{target}' uses an unsupported scheme. " +
            "Supported: http(s)://, mailto:, file://, internal #Sheet!Range " +
            "(decision I13).", nameof(target));
    }

    public string? GetFormula()
    {
        _workbook.ThrowIfDisposed();
        if (_underlying.CellType != CellType.Formula) return null;
        var body = _underlying.CellFormula;
        return body is null ? null : "=" + body;
    }

    public void Clear()
    {
        _workbook.ThrowIfDisposed();
        _underlying.SetBlank();
    }

    public string GetString()
    {
        _workbook.ThrowIfDisposed();
        // Per design §7.10:
        //   Empty -> ""; String -> stored verbatim; Number -> invariant string;
        //   Bool -> "TRUE" / "FALSE" (invariant); Formula -> cached result;
        //   Error -> error code text.
        switch (_underlying.CellType)
        {
            case CellType.Blank: return string.Empty;
            case CellType.String: return _underlying.StringCellValue ?? string.Empty;
            case CellType.Boolean: return _underlying.BooleanCellValue ? "TRUE" : "FALSE";
            case CellType.Numeric:
                return _underlying.NumericCellValue.ToString("G17", CultureInfo.InvariantCulture);
            case CellType.Error:
                return FormulaError.ForInt(_underlying.ErrorCellValue).String;
            case CellType.Formula:
                return GetFormulaCachedAsString();
            default: return string.Empty;
        }
    }

    private string GetFormulaCachedAsString()
    {
        switch (_underlying.CachedFormulaResultType)
        {
            case CellType.String: return _underlying.StringCellValue ?? string.Empty;
            case CellType.Boolean: return _underlying.BooleanCellValue ? "TRUE" : "FALSE";
            case CellType.Numeric:
                return _underlying.NumericCellValue.ToString("G17", CultureInfo.InvariantCulture);
            case CellType.Error:
                return FormulaError.ForInt(_underlying.ErrorCellValue).String;
            default: return string.Empty;
        }
    }

    public double? GetNumber()
    {
        _workbook.ThrowIfDisposed();
        switch (_underlying.CellType)
        {
            case CellType.Numeric: return _underlying.NumericCellValue;
            case CellType.Boolean: return _underlying.BooleanCellValue ? 1.0 : 0.0;
            case CellType.Formula:
                return _underlying.CachedFormulaResultType == CellType.Numeric
                    ? _underlying.NumericCellValue
                    : (double?)null;
            default: return null;
        }
    }

    public bool? GetBool()
    {
        _workbook.ThrowIfDisposed();
        return _underlying.CellType switch
        {
            CellType.Boolean => _underlying.BooleanCellValue,
            CellType.Formula when _underlying.CachedFormulaResultType == CellType.Boolean
                => _underlying.BooleanCellValue,
            _ => null,
        };
    }

    public DateTime? GetDate()
    {
        _workbook.ThrowIfDisposed();
        if (_underlying.CellType != CellType.Numeric) return null;
        if (!DateUtil.IsCellDateFormatted(_underlying)) return null;
        var dt = _underlying.DateCellValue;
        return dt is null ? null : DateTime.SpecifyKind(dt.Value, DateTimeKind.Unspecified);
    }

    public DateOnly? GetDateOnly()
    {
        var dt = GetDate();
        return dt is null ? null : DateOnly.FromDateTime(dt.Value);
    }

    public TimeOnly? GetTime()
    {
        // §7.9: accepts any numeric cell; returns null when out of TimeOnly range.
        var num = GetNumber();
        if (num is null) return null;
        if (num.Value < 0.0 || num.Value >= 1.0) return null;
        return TimeOnly.FromTimeSpan(TimeSpan.FromDays(num.Value));
    }

    public TimeSpan? GetDuration()
    {
        var num = GetNumber();
        return num is null ? null : TimeSpan.FromDays(num.Value);
    }

    public CellError? GetError()
    {
        _workbook.ThrowIfDisposed();
        var type = _underlying.CellType;
        if (type == CellType.Error)
        {
            return MapNpoiErrorCode(_underlying.ErrorCellValue);
        }
        if (type == CellType.Formula && _underlying.CachedFormulaResultType == CellType.Error)
        {
            return MapNpoiErrorCode(_underlying.ErrorCellValue);
        }
        return null;
    }

    private static CellError MapNpoiErrorCode(byte code) => code switch
    {
        0x00 => CellError.Null,         // #NULL!
        0x07 => CellError.DivByZero,    // #DIV/0!
        0x0F => CellError.Value,        // #VALUE!
        0x17 => CellError.Ref,          // #REF!
        0x1D => CellError.Name,         // #NAME?
        0x24 => CellError.Num,          // #NUM!
        0x2A => CellError.NotAvailable, // #N/A
        0x2B => CellError.GettingData,  // #GETTING_DATA
        _    => CellError.Value,        // unknown — best fallback to #VALUE!
    };

    public ICell Style(CellStyle style)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(style);

        // Merge: non-null properties in `style` overwrite the cell's
        // current style on that axis; null means "inherit existing."
        // CellStyle.Default (all-null) is a no-op.
        var current = ReadCurrentStyle();
        var merged = Merge(current, style);

        var npoiStyle = _workbook.StylePool.GetOrCreate(merged);
        _underlying.CellStyle = npoiStyle;
        return this;
    }

    public ICell NumberFormat(string format)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(format);
        return Style(new CellStyle { NumberFormat = format });
    }

    public CellStyle GetStyle()
    {
        _workbook.ThrowIfDisposed();
        return ReadCurrentStyle();
    }

    private CellStyle ReadCurrentStyle()
    {
        var ns = _underlying.CellStyle;
        // Index 0 is the workbook default — treat as "no style applied".
        if (ns is null || ns.Index == 0) return CellStyle.Default;
        return CellStylePool.ReadFromNpoi(ns);
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
        NumberFormat = overlay.NumberFormat ?? existing.NumberFormat,
        HorizontalAlignment = overlay.HorizontalAlignment ?? existing.HorizontalAlignment,
        VerticalAlignment = overlay.VerticalAlignment ?? existing.VerticalAlignment,
        WrapText = overlay.WrapText ?? existing.WrapText,
        Borders = overlay.Borders ?? existing.Borders,
    };

    public XSSFCell Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _underlying; }
    }

    /// <summary>
    /// Applies <paramref name="defaultStyle"/> to the cell when it currently
    /// carries no explicit style. Per decision I-18: a user-set style is
    /// preserved. The workbook-default style has index 0; any explicit
    /// style has a higher index.
    /// </summary>
    private void ApplyDefaultStyleIfUnstyled(ICellStyle defaultStyle)
    {
        var current = _underlying.CellStyle;
        if (current is null || current.Index == 0)
        {
            _underlying.CellStyle = defaultStyle;
        }
    }
}
