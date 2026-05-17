// Exception hierarchy per design §6.13.

using System;

namespace NetXlsx;

/// <summary>Base exception for all NetXlsx errors.</summary>
public class WorkbookException : Exception
{
    /// <summary>Creates a new <see cref="WorkbookException"/>.</summary>
    public WorkbookException() { }
    /// <summary>Creates a new <see cref="WorkbookException"/> with the given message.</summary>
    public WorkbookException(string message) : base(message) { }
    /// <summary>Creates a new <see cref="WorkbookException"/> with the given message and inner exception.</summary>
    public WorkbookException(string message, Exception innerException) : base(message, innerException) { }
}

/// <summary>
/// Thrown when a cell address string fails to parse (e.g. <c>"foo"</c>,
/// <c>"Sheet1!A1"</c>, <c>"A:A"</c> in a single-cell context). See
/// <c>docs/design.md §6.10</c> for the accepted grammar.
/// </summary>
public sealed class InvalidCellAddressException : WorkbookException
{
    /// <summary>The string that failed to parse.</summary>
    public string Input { get; }

    /// <summary>Creates a new <see cref="InvalidCellAddressException"/>.</summary>
    public InvalidCellAddressException(string input, string reason)
        : base($"Invalid cell address '{input}': {reason}")
    {
        Input = input;
    }
}

/// <summary>
/// Thrown when a sheet name violates Excel's rules (length 1..31, no
/// <c>\ / ? * [ ]</c>, must be unique within the workbook).
/// </summary>
public sealed class SheetNameException : WorkbookException
{
    /// <summary>The offending sheet name.</summary>
    public string Name { get; }

    /// <summary>Creates a new <see cref="SheetNameException"/>.</summary>
    public SheetNameException(string name, string reason)
        : base($"Invalid sheet name '{name}': {reason}")
    {
        Name = name;
    }
}

/// <summary>
/// Thrown when opening a file that is not a valid <c>.xlsx</c> workbook —
/// e.g., an <c>.xls</c> binary, a corrupt OPC package, or arbitrary
/// non-OOXML content. Wraps the underlying NPOI exception when available.
/// </summary>
public sealed class MalformedFileException : WorkbookException
{
    /// <summary>Creates a new <see cref="MalformedFileException"/>.</summary>
    public MalformedFileException(string message) : base(message) { }
    /// <summary>Creates a new <see cref="MalformedFileException"/> with inner exception.</summary>
    public MalformedFileException(string message, Exception innerException) : base(message, innerException) { }
}

/// <summary>
/// Thrown when a write would exceed an Excel hard limit — cell text
/// over 32,767 characters, row index past 1,048,576, column past XFD
/// (16,384) — per decision #37 / §7.6. Fail loud rather than silently
/// truncate or wrap.
/// </summary>
public sealed class ResourceLimitExceededException : WorkbookException
{
    /// <summary>The limit's name (e.g. <c>"cell text length"</c>).</summary>
    public string LimitName { get; }
    /// <summary>The maximum allowed value.</summary>
    public long Limit { get; }
    /// <summary>The actual value that triggered the failure.</summary>
    public long Actual { get; }

    /// <summary>Creates a new <see cref="ResourceLimitExceededException"/>.</summary>
    public ResourceLimitExceededException(string limitName, long limit, long actual)
        : base($"Excel hard limit exceeded — {limitName}: {actual} exceeds maximum {limit}.")
    {
        LimitName = limitName;
        Limit = limit;
        Actual = actual;
    }
}

/// <summary>
/// Thrown by <see cref="IColumn.AutoSize"/> when the runtime environment
/// cannot supply font metrics — typically headless Linux without
/// <c>libgdiplus</c> and a fallback font installed (design decision I3).
/// Message includes installation guidance for common distributions.
/// </summary>
public sealed class MissingFontException : WorkbookException
{
    /// <summary>Creates a new <see cref="MissingFontException"/>.</summary>
    public MissingFontException()
        : base(BuildMessage()) { }
    /// <summary>Creates a new <see cref="MissingFontException"/> with a custom message.</summary>
    public MissingFontException(string message) : base(message) { }
    /// <summary>Creates a new <see cref="MissingFontException"/> wrapping the underlying font failure.</summary>
    public MissingFontException(Exception innerException)
        : base(BuildMessage(), innerException) { }
    /// <summary>Creates a new <see cref="MissingFontException"/> with message and inner exception.</summary>
    public MissingFontException(string message, Exception innerException)
        : base(message, innerException) { }

    private static string BuildMessage() =>
        "IColumn.AutoSize() requires font metrics that are not available in this environment. "
        + "On headless Linux, install a font package and (if needed) libgdiplus: "
        + "Debian/Ubuntu — 'apt-get install -y fontconfig fonts-dejavu-core libgdiplus'; "
        + "Alpine — 'apk add fontconfig ttf-dejavu libgdiplus'. "
        + "Alternatively, set column widths explicitly via IColumn.Width(double) (design decision I3).";
}
