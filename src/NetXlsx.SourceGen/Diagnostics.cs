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
    /// A <c>[Column(Format = "...")]</c> string fails the structural smoke
    /// check. This is intentionally a narrow check (empty / control chars /
    /// unbalanced brackets / no recognized format-character) — not full
    /// Excel format-grammar validation, which Excel itself does at render
    /// time.
    /// </summary>
    public static readonly DiagnosticDescriptor MalformedFormatString = new(
        id: "NXLS0003",
        title: "[Column] Format string fails structural smoke check",
        messageFormat: "Property '{0}.{1}' has [Column(Format = \"{2}\")] which fails the structural smoke check ({3}). Full Excel format-grammar validation is not performed; Excel reports format errors at render time.",
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
    /// A property has a type the built-in converter set does not support.
    /// Nullable&lt;T&gt; and other unmodeled types are handled via a custom
    /// converter (<c>[Column(ConverterType = ...)]</c>, decision I-58).
    /// </summary>
    public static readonly DiagnosticDescriptor UnsupportedPropertyType = new(
        id: "NXLS0006",
        title: "[Worksheet] property type has no built-in converter",
        messageFormat: "Property '{0}.{1}' has type '{2}', which has no built-in converter. Supported types: string, bool, byte/short/int/long (and unsigned), float/double/decimal, DateTime, DateOnly, TimeOnly, TimeSpan. Other types (including Nullable<T>) need a custom converter via [Column(ConverterType = typeof(...))].",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);
}
