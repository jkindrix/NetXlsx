// XssfCell — value setters and getters. SetX/GetX for every scalar
// type the public ICell surface supports, plus the formula + error
// surface. Core class structure is in XssfCell.cs.

using System;
using System.Globalization;
using NPOI.SS.UserModel;

namespace NetXlsx;

internal sealed partial class XssfCell
{
    // ---- Setters --------------------------------------------------------

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

    // ---- Getters --------------------------------------------------------

    public string? GetFormula()
    {
        _workbook.ThrowIfDisposed();
        if (_underlying.CellType != CellType.Formula) return null;
        var body = _underlying.CellFormula;
        return body is null ? null : "=" + body;
    }

    public string GetString()
    {
        _workbook.ThrowIfDisposed();
        // Per design §7.10:
        //   Empty -> ""; String -> stored verbatim;
        //   Number (no format) -> invariant G17;
        //   Date -> display-formatted via DisplayCulture (uses NPOI's
        //           DataFormatter so the cell's format string drives output);
        //   Bool -> "TRUE" / "FALSE" (invariant; never localized);
        //   Formula -> cached result; Error -> error code text.
        switch (_underlying.CellType)
        {
            case CellType.Blank: return string.Empty;
            case CellType.String: return _underlying.StringCellValue ?? string.Empty;
            case CellType.Boolean: return _underlying.BooleanCellValue ? "TRUE" : "FALSE";
            case CellType.Numeric:
                if (DateUtil.IsCellDateFormatted(_underlying))
                {
                    var formatter = new DataFormatter(_workbook.Options.DisplayCulture);
                    return formatter.FormatCellValue(_underlying);
                }
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
}
