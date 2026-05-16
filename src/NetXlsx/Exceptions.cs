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
