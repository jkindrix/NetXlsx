// Value model captured by the incremental pipeline. Must be value-equal
// so Roslyn can cache pipeline output — that's the point of
// IIncrementalGenerator. Avoids `init` / `required` because this assembly
// targets netstandard2.0 (Roslyn analyzer host requirement) and lacks the
// supporting BCL types.

using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using Microsoft.CodeAnalysis;

namespace NetXlsx.SourceGen;

internal sealed class WorksheetModel : IEquatable<WorksheetModel>
{
    public string FullyQualifiedName { get; }
    public string Namespace { get; }
    public string TypeName { get; }
    public TypeKindLite Kind { get; }
    public string Visibility { get; }
    public bool IsPartial { get; }
    public bool HasDesignatedConstructor { get; }
    /// <summary>
    /// Record primary-constructor parameters, declaration order. Empty for
    /// non-records and records without a parameter list. When non-empty,
    /// the generated <c>ReadRows</c> constructs through this constructor
    /// instead of an object initializer (positional records have no
    /// parameterless constructor — object-initializer emit is CS7036).
    /// </summary>
    public EquatableArray<PrimaryCtorParam> PrimaryCtorParameters { get; }
    public EquatableArray<WorksheetProperty> Properties { get; }
    public EquatableArray<DiagnosticInfo> EarlyDiagnostics { get; }

    public WorksheetModel(
        string fullyQualifiedName, string @namespace, string typeName,
        TypeKindLite kind, string visibility, bool isPartial, bool hasCtor,
        EquatableArray<PrimaryCtorParam> primaryCtorParameters,
        EquatableArray<WorksheetProperty> properties,
        EquatableArray<DiagnosticInfo> earlyDiagnostics)
    {
        FullyQualifiedName = fullyQualifiedName;
        Namespace = @namespace;
        TypeName = typeName;
        Kind = kind;
        Visibility = visibility;
        IsPartial = isPartial;
        HasDesignatedConstructor = hasCtor;
        PrimaryCtorParameters = primaryCtorParameters;
        Properties = properties;
        EarlyDiagnostics = earlyDiagnostics;
    }

    public bool Equals(WorksheetModel? other) =>
        other is not null
        && FullyQualifiedName == other.FullyQualifiedName
        && Namespace == other.Namespace
        && TypeName == other.TypeName
        && Kind == other.Kind
        && Visibility == other.Visibility
        && IsPartial == other.IsPartial
        && HasDesignatedConstructor == other.HasDesignatedConstructor
        && PrimaryCtorParameters.Equals(other.PrimaryCtorParameters)
        && Properties.Equals(other.Properties)
        && EarlyDiagnostics.Equals(other.EarlyDiagnostics);

    public override bool Equals(object? obj) => Equals(obj as WorksheetModel);

    public override int GetHashCode()
    {
        unchecked
        {
            int h = 17;
            h = h * 31 + FullyQualifiedName.GetHashCode();
            h = h * 31 + Namespace.GetHashCode();
            h = h * 31 + TypeName.GetHashCode();
            h = h * 31 + (int)Kind;
            h = h * 31 + Visibility.GetHashCode();
            h = h * 31 + IsPartial.GetHashCode();
            h = h * 31 + HasDesignatedConstructor.GetHashCode();
            h = h * 31 + PrimaryCtorParameters.GetHashCode();
            h = h * 31 + Properties.GetHashCode();
            h = h * 31 + EarlyDiagnostics.GetHashCode();
            return h;
        }
    }
}

internal sealed class WorksheetProperty : IEquatable<WorksheetProperty>
{
    public string Name { get; }
    public string FullTypeName { get; }
    public SpecialType UnderlyingSpecialType { get; }
    public bool IsNullable { get; }
    public ColumnMapping? Column { get; }
    public bool IsIgnored { get; }
    public PropertyLocation Location { get; }
    public bool TypeIsSupported { get; }
    /// <summary>True when the property has a public setter or init accessor (object-initializer assignable).</summary>
    public bool IsSettable { get; }
    /// <summary>True when the property is declared <c>required</c>.</summary>
    public bool IsRequired { get; }
    /// <summary>Fully-qualified converter type, or null when no converter is configured.</summary>
    public string? ConverterTypeFullName { get; }

    public WorksheetProperty(string name, string fullTypeName, SpecialType underlyingSpecialType,
        bool isNullable, ColumnMapping? column, bool isIgnored, PropertyLocation location, bool typeIsSupported,
        bool isSettable, bool isRequired, string? converterTypeFullName)
    {
        Name = name;
        FullTypeName = fullTypeName;
        UnderlyingSpecialType = underlyingSpecialType;
        IsNullable = isNullable;
        Column = column;
        IsIgnored = isIgnored;
        Location = location;
        TypeIsSupported = typeIsSupported;
        IsSettable = isSettable;
        IsRequired = isRequired;
        ConverterTypeFullName = converterTypeFullName;
    }

    // Location is deliberately EXCLUDED from equality/hash. It carries
    // file/line/column, which shift on any trivia edit (whitespace, a
    // comment added above the property). Including it in the cache key
    // would bust incrementality on edits that don't change generated
    // output. The accepted trade-off: after a pure-trivia move the cached
    // diagnostic squiggle for this property may point at the old location
    // until some output-affecting change triggers a regen.
    public bool Equals(WorksheetProperty? other) =>
        other is not null
        && Name == other.Name
        && FullTypeName == other.FullTypeName
        && UnderlyingSpecialType == other.UnderlyingSpecialType
        && IsNullable == other.IsNullable
        && Equals(Column, other.Column)
        && IsIgnored == other.IsIgnored
        && TypeIsSupported == other.TypeIsSupported
        && IsSettable == other.IsSettable
        && IsRequired == other.IsRequired
        && ConverterTypeFullName == other.ConverterTypeFullName;

    public override bool Equals(object? obj) => Equals(obj as WorksheetProperty);

    public override int GetHashCode()
    {
        unchecked
        {
            int h = 17;
            h = h * 31 + Name.GetHashCode();
            h = h * 31 + FullTypeName.GetHashCode();
            h = h * 31 + (int)UnderlyingSpecialType;
            h = h * 31 + IsNullable.GetHashCode();
            h = h * 31 + (Column?.GetHashCode() ?? 0);
            h = h * 31 + IsIgnored.GetHashCode();
            h = h * 31 + TypeIsSupported.GetHashCode();
            h = h * 31 + IsSettable.GetHashCode();
            h = h * 31 + IsRequired.GetHashCode();
            h = h * 31 + (ConverterTypeFullName?.GetHashCode() ?? 0);
            return h;
        }
    }
}

internal sealed class ColumnMapping : IEquatable<ColumnMapping>
{
    public string HeaderName { get; }
    public int Order { get; }
    public string? Format { get; }

    public ColumnMapping(string headerName, int order, string? format)
    {
        HeaderName = headerName;
        Order = order;
        Format = format;
    }

    public bool Equals(ColumnMapping? other) =>
        other is not null
        && HeaderName == other.HeaderName
        && Order == other.Order
        && Format == other.Format;

    public override bool Equals(object? obj) => Equals(obj as ColumnMapping);

    public override int GetHashCode()
    {
        unchecked
        {
            int h = 17;
            h = h * 31 + HeaderName.GetHashCode();
            h = h * 31 + Order;
            h = h * 31 + (Format?.GetHashCode() ?? 0);
            return h;
        }
    }
}

/// <summary>
/// One record primary-constructor parameter, captured for ReadRows
/// construction (positional records expose mapped values as ctor params,
/// not initializer-assignable properties).
/// </summary>
internal sealed class PrimaryCtorParam : IEquatable<PrimaryCtorParam>
{
    public string Name { get; }
    public bool HasDefaultValue { get; }
    public PropertyLocation Location { get; }

    public PrimaryCtorParam(string name, bool hasDefaultValue, PropertyLocation location)
    {
        Name = name;
        HasDefaultValue = hasDefaultValue;
        Location = location;
    }

    // Location is excluded from equality/hash for the same incrementality
    // reason as WorksheetProperty.Location (see comment there).
    public bool Equals(PrimaryCtorParam? other) =>
        other is not null
        && Name == other.Name
        && HasDefaultValue == other.HasDefaultValue;

    public override bool Equals(object? obj) => Equals(obj as PrimaryCtorParam);

    public override int GetHashCode()
    {
        unchecked
        {
            int h = 17;
            h = h * 31 + Name.GetHashCode();
            h = h * 31 + HasDefaultValue.GetHashCode();
            return h;
        }
    }
}

internal enum TypeKindLite
{
    Class,
    Record,
    RecordStruct,
    Struct,
}

internal sealed class PropertyLocation : IEquatable<PropertyLocation>
{
    public string FilePath { get; }
    public int Line { get; }
    public int Column { get; }
    public int Length { get; }

    public PropertyLocation(string filePath, int line, int column, int length)
    {
        FilePath = filePath;
        Line = line;
        Column = column;
        Length = length;
    }

    public Location ToRoslynLocation()
    {
        var linePosition = new Microsoft.CodeAnalysis.Text.LinePosition(Line, Column);
        var span = new Microsoft.CodeAnalysis.Text.LinePositionSpan(linePosition,
            new Microsoft.CodeAnalysis.Text.LinePosition(Line, Column + Length));
        return Location.Create(FilePath, default, span);
    }

    public bool Equals(PropertyLocation? other) =>
        other is not null
        && FilePath == other.FilePath
        && Line == other.Line
        && Column == other.Column
        && Length == other.Length;

    public override bool Equals(object? obj) => Equals(obj as PropertyLocation);

    public override int GetHashCode()
    {
        unchecked
        {
            int h = 17;
            h = h * 31 + FilePath.GetHashCode();
            h = h * 31 + Line;
            h = h * 31 + Column;
            h = h * 31 + Length;
            return h;
        }
    }
}

internal sealed class DiagnosticInfo : IEquatable<DiagnosticInfo>
{
    public DiagnosticDescriptor Descriptor { get; }
    public EquatableArray<string> MessageArgs { get; }
    public PropertyLocation? Location { get; }

    public DiagnosticInfo(DiagnosticDescriptor descriptor, EquatableArray<string> messageArgs, PropertyLocation? location)
    {
        Descriptor = descriptor;
        MessageArgs = messageArgs;
        Location = location;
    }

    // Descriptor reference-equality is safe here: the only descriptor
    // instances in the pipeline are the static singletons on `Diagnostics`,
    // so two `DiagnosticInfo` values referring to the "same" diagnostic
    // share the same reference.
    public bool Equals(DiagnosticInfo? other) =>
        other is not null
        && ReferenceEquals(Descriptor, other.Descriptor)
        && MessageArgs.Equals(other.MessageArgs)
        && Equals(Location, other.Location);

    public override bool Equals(object? obj) => Equals(obj as DiagnosticInfo);

    public override int GetHashCode()
    {
        unchecked
        {
            int h = 17;
            h = h * 31 + Descriptor.Id.GetHashCode();
            h = h * 31 + MessageArgs.GetHashCode();
            h = h * 31 + (Location?.GetHashCode() ?? 0);
            return h;
        }
    }
}

/// <summary>
/// A struct wrapper over <see cref="ImmutableArray{T}"/> that implements
/// structural equality — required for IIncrementalGenerator caching.
/// (ImmutableArray itself uses reference equality on the backing array.)
/// </summary>
internal readonly struct EquatableArray<T> : IEquatable<EquatableArray<T>>
    where T : IEquatable<T>
{
    public ImmutableArray<T> Array { get; }
    public int Length => Array.IsDefault ? 0 : Array.Length;
    public T this[int i] => Array[i];

    public EquatableArray(ImmutableArray<T> array) { Array = array; }
    public static EquatableArray<T> Empty => new(ImmutableArray<T>.Empty);

    public bool Equals(EquatableArray<T> other)
    {
        if (Array.IsDefault) return other.Array.IsDefault;
        if (other.Array.IsDefault) return false;
        if (Array.Length != other.Array.Length) return false;
        for (int i = 0; i < Array.Length; i++)
            if (!Array[i].Equals(other.Array[i])) return false;
        return true;
    }

    public override bool Equals(object? obj) => obj is EquatableArray<T> e && Equals(e);

    public override int GetHashCode()
    {
        if (Array.IsDefault) return 0;
        unchecked
        {
            int hash = 17;
            foreach (var item in Array) hash = hash * 31 + (item?.GetHashCode() ?? 0);
            return hash;
        }
    }

    public IEnumerator<T> GetEnumerator() => ((IEnumerable<T>)Array).GetEnumerator();
}
