// Public marker attributes consumed by the NetXlsx source generator.
// See docs/design.md §6.9 (typed mapping) and §6.12 (diagnostic catalog).

using System;

namespace NetXlsx;

/// <summary>
/// Marks a record / class as a worksheet row schema. The NetXlsx source
/// generator emits typed read/write extension methods on <c>ISheet</c> for
/// every type annotated with this attribute.
/// </summary>
/// <remarks>
/// <para>
/// The attributed type must be <c>partial</c> (decision <c>NXLS0005</c>)
/// and must expose a designated constructor — either a public parameterless
/// constructor or a primary constructor on a record
/// (decision <c>NXLS0002</c>).
/// </para>
/// <para>
/// <b>Scoping (decision I5):</b> the generator scans only the *current
/// compilation*. Types annotated with <c>[Worksheet]</c> in a
/// <em>referenced assembly</em> are <b>invisible</b> to the generator
/// running in the consuming project — the extension methods are not emitted,
/// and calls to them do not compile. Common failure mode: a consumer puts
/// shared records in a "Domain" library expecting the extensions to work in
/// the calling app. They do not. The fix is to add the <c>NetXlsx</c>
/// package to the Domain library too, so the generator runs against that
/// compilation. This matches the scoping rule used by
/// <c>System.Text.Json</c>'s source generator.
/// </para>
/// </remarks>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct, AllowMultiple = false, Inherited = false)]
public sealed class WorksheetAttribute : Attribute
{
    /// <summary>
    /// Visibility of the generated extension class. Defaults to
    /// <see cref="WorksheetVisibility.Internal"/>.
    /// </summary>
    public WorksheetVisibility Visibility { get; init; } = WorksheetVisibility.Internal;
}

/// <summary>
/// Visibility of a source-generated worksheet extensions class.
/// </summary>
public enum WorksheetVisibility
{
    /// <summary>Generated class is <c>internal</c> (default).</summary>
    Internal,
    /// <summary>Generated class is <c>public</c> — opt in.</summary>
    Public,
}

/// <summary>
/// Maps a property on a <c>[Worksheet]</c>-annotated type to a column in the
/// sheet by header name, column index, format, or order.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public sealed class ColumnAttribute : Attribute
{
    /// <summary>The column's header text. Matched case-insensitively (ordinal).</summary>
    public string Name { get; }

    /// <summary>
    /// Optional explicit column ordering when writing. If unset, properties
    /// are written in source declaration order. If multiple properties
    /// specify the same <see cref="Order"/>, the generator emits diagnostic
    /// <c>NXLS0001</c>.
    /// </summary>
    public int Order { get; init; } = -1;

    /// <summary>
    /// Excel number-format string applied to the cell on write
    /// (e.g. <c>$#,##0.00</c>, <c>yyyy-mm-dd</c>). Pass-through bytes —
    /// the generator does not localize.
    /// </summary>
    public string? Format { get; init; }

    /// <summary>
    /// Creates a column mapping bound to the supplied header name.
    /// </summary>
    public ColumnAttribute(string name)
    {
        Name = name;
    }
}

/// <summary>
/// Excludes a property from typed row mapping. The property is neither read
/// nor written and produces no diagnostic.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public sealed class IgnoreAttribute : Attribute
{
}
