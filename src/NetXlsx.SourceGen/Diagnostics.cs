// NXLS0001-NXLS0006 diagnostic catalog.
// See docs/design.md §6.12.

using Microsoft.CodeAnalysis;

namespace NetXlsx.SourceGen;

internal static class Diagnostics
{
    private const string Category = "NetXlsx";

    /// <summary>
    /// Two or more <c>[Column]</c> attributes specify the same explicit
    /// <c>Order</c> value, producing ambiguous write ordering.
    /// </summary>
    public static readonly DiagnosticDescriptor DuplicateColumnOrder = new(
        id: "NXLS0001",
        title: "[Worksheet] type has duplicate [Column] Order values",
        messageFormat: "Type '{0}' has multiple [Column] attributes with Order = {1}: {2}. Each explicit Order must be unique.",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    /// <summary>
    /// The type has no public parameterless constructor and no record
    /// primary constructor.
    /// </summary>
    public static readonly DiagnosticDescriptor MissingDesignatedConstructor = new(
        id: "NXLS0002",
        title: "[Worksheet] type has no designated constructor",
        messageFormat: "Type '{0}' has no public parameterless constructor and no record primary constructor. Add one or convert the type to a record with a primary constructor.",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    /// <summary>
    /// A <c>[Column(Format = "...")]</c> string failed the format-string
    /// smoke check (e.g. unmatched bracket, suspicious characters).
    /// </summary>
    public static readonly DiagnosticDescriptor InvalidFormatString = new(
        id: "NXLS0003",
        title: "[Column] Format string failed smoke check",
        messageFormat: "Property '{0}.{1}' has [Column(Format = \"{2}\")] which fails the format-string smoke check: {3}",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    /// <summary>
    /// A property on a <c>[Worksheet]</c> type carries neither
    /// <c>[Column]</c> nor <c>[Ignore]</c> — silent skip would be ambiguous.
    /// </summary>
    public static readonly DiagnosticDescriptor UnmappedProperty = new(
        id: "NXLS0004",
        title: "[Worksheet] property is neither mapped nor ignored",
        messageFormat: "Property '{0}.{1}' on a [Worksheet] type has neither [Column] nor [Ignore]. Mark it explicitly to avoid ambiguous behavior.",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    /// <summary>
    /// A <c>[Worksheet]</c> type must be declared <c>partial</c> so the
    /// generator can emit nested helpers alongside it in later versions.
    /// </summary>
    public static readonly DiagnosticDescriptor TypeNotPartial = new(
        id: "NXLS0005",
        title: "[Worksheet] type must be partial",
        messageFormat: "Type '{0}' is annotated with [Worksheet] but is not declared 'partial'. Add the 'partial' modifier.",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    /// <summary>
    /// A property has a type the v1.0 built-in converter set does not
    /// support (custom converters arrive in v1.1).
    /// </summary>
    public static readonly DiagnosticDescriptor UnsupportedPropertyType = new(
        id: "NXLS0006",
        title: "[Worksheet] property type has no built-in converter",
        messageFormat: "Property '{0}.{1}' has type '{2}', which v1.0 has no built-in converter for. Supported types: string, bool, byte/short/int/long (and unsigned), float/double/decimal, DateTime, DateOnly, TimeOnly, TimeSpan, and Nullable<T> over any of these.",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);
}
