// NetXlsx [Worksheet] source generator.
// See docs/design.md §6.9 (typed mapping) and §6.12 (diagnostic catalog).

using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;

namespace NetXlsx.SourceGen;

/// <summary>
/// IIncrementalGenerator that scans the current compilation for types
/// annotated with <c>[NetXlsx.Worksheet]</c> and emits typed
/// read/write extension methods on <c>NetXlsx.ISheet</c>.
/// </summary>
/// <remarks>
/// Scoping: only types defined in the current compilation are scanned.
/// <c>[Worksheet]</c> types in *referenced* assemblies are ignored
/// (decision I5). Each assembly that defines <c>[Worksheet]</c> types
/// must add the <c>NetXlsx</c> package so the generator runs there.
/// </remarks>
[Generator(LanguageNames.CSharp)]
public sealed class WorksheetGenerator : IIncrementalGenerator
{
    private const string WorksheetAttributeFullName = "NetXlsx.WorksheetAttribute";
    private const string ColumnAttributeFullName = "NetXlsx.ColumnAttribute";
    private const string IgnoreAttributeFullName = "NetXlsx.IgnoreAttribute";

    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var models = context.SyntaxProvider
            .ForAttributeWithMetadataName(
                WorksheetAttributeFullName,
                predicate: static (node, _) => node is TypeDeclarationSyntax,
                transform: static (ctx, ct) => Transform(ctx, ct))
            .Where(static m => m is not null)!;

        context.RegisterSourceOutput(models, static (spc, model) =>
        {
            if (model is null) return;

            // Replay diagnostics captured during transform.
            foreach (var d in model.EarlyDiagnostics)
            {
                spc.ReportDiagnostic(BuildDiagnostic(d));
            }

            // Property-level diagnostics (computed at emit time so we don't
            // bloat the early-diagnostics list).
            EmitPropertyDiagnostics(spc, model);

            // If the type has fatal early diagnostics (not partial, no ctor),
            // skip code emission — the file would not compile anyway.
            if (HasFatalEarlyDiagnostic(model)) return;

            var source = EmitExtensionsSource(model);
            var hintName = $"{Sanitize(model.FullyQualifiedName)}.g.cs";
            spc.AddSource(hintName, SourceText.From(source, Encoding.UTF8));
        });
    }

    // ------------------------------------------------------------------
    // Transform: Roslyn semantic model -> WorksheetModel (data-only)
    // ------------------------------------------------------------------

    private static WorksheetModel? Transform(GeneratorAttributeSyntaxContext ctx, CancellationToken ct)
    {
        if (ctx.TargetSymbol is not INamedTypeSymbol type) return null;
        var syntax = (TypeDeclarationSyntax)ctx.TargetNode;
        ct.ThrowIfCancellationRequested();

        var earlyDiagnostics = ImmutableArray.CreateBuilder<DiagnosticInfo>();

        var isPartial = syntax.Modifiers.Any(m => m.IsKind(SyntaxKind.PartialKeyword));
        if (!isPartial)
        {
            earlyDiagnostics.Add(new DiagnosticInfo(
                Diagnostics.TypeNotPartial,
                new EquatableArray<string>(ImmutableArray.Create(type.ToDisplayString())),
                LocationFrom(syntax.Identifier.GetLocation())));
        }

        var hasCtor = HasDesignatedConstructor(type, syntax);
        if (!hasCtor)
        {
            earlyDiagnostics.Add(new DiagnosticInfo(
                Diagnostics.MissingDesignatedConstructor,
                new EquatableArray<string>(ImmutableArray.Create(type.ToDisplayString())),
                LocationFrom(syntax.Identifier.GetLocation())));
        }

        var visibility = GetVisibility(ctx);

        var properties = BuildProperties(type, ctx.SemanticModel.Compilation, ct);

        // Type-level: duplicate explicit Order across [Column] attributes
        var orderGroups = properties
            .Where(p => p.Column is not null && p.Column.Order >= 0)
            .GroupBy(p => p.Column!.Order)
            .Where(g => g.Count() > 1)
            .ToList();
        foreach (var grp in orderGroups)
        {
            var names = string.Join(", ", grp.Select(p => p.Name));
            earlyDiagnostics.Add(new DiagnosticInfo(
                Diagnostics.DuplicateColumnOrder,
                new EquatableArray<string>(ImmutableArray.Create(
                    type.ToDisplayString(), grp.Key.ToString(System.Globalization.CultureInfo.InvariantCulture), names)),
                LocationFrom(syntax.Identifier.GetLocation())));
        }

        return new WorksheetModel(
            fullyQualifiedName: type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
            @namespace: type.ContainingNamespace?.IsGlobalNamespace == true
                ? string.Empty
                : type.ContainingNamespace!.ToDisplayString(),
            typeName: type.Name,
            kind: ClassifyKind(type),
            visibility: visibility,
            isPartial: isPartial,
            hasCtor: hasCtor,
            properties: new EquatableArray<WorksheetProperty>(properties.ToImmutableArray()),
            earlyDiagnostics: new EquatableArray<DiagnosticInfo>(earlyDiagnostics.ToImmutable()));
    }

    private static List<WorksheetProperty> BuildProperties(
        INamedTypeSymbol type, Compilation compilation, CancellationToken ct)
    {
        // Walk the type and its base chain (excluding System.Object) so a
        // [Worksheet] type that inherits mappable public properties picks
        // them up instead of silently dropping them. Base-most first, so
        // inherited columns lead the derived type's own (matching how the
        // properties are declared top-to-bottom in the hierarchy). A
        // property re-declared in a derived type (override or `new`) keeps
        // its base column position but takes the most-derived metadata.
        var chain = new List<INamedTypeSymbol>();
        for (var t = type; t is not null && t.SpecialType != SpecialType.System_Object; t = t.BaseType)
            chain.Add(t);
        chain.Reverse();

        var result = new List<WorksheetProperty>();
        var indexByName = new Dictionary<string, int>(StringComparer.Ordinal);

        foreach (var declaringType in chain)
        foreach (var member in declaringType.GetMembers().OfType<IPropertySymbol>())
        {
            ct.ThrowIfCancellationRequested();
            if (member.DeclaredAccessibility != Accessibility.Public) continue;
            if (member.IsStatic) continue;
            if (member.IsIndexer) continue;

            var attrs = member.GetAttributes();
            var ignored = attrs.Any(a => a.AttributeClass?.ToDisplayString() == IgnoreAttributeFullName);
            var colAttr = attrs.FirstOrDefault(a => a.AttributeClass?.ToDisplayString() == ColumnAttributeFullName);

            ColumnMapping? column = null;
            string? converterTypeName = null;
            if (colAttr is not null && colAttr.ConstructorArguments.Length > 0)
            {
                var headerName = colAttr.ConstructorArguments[0].Value as string ?? member.Name;
                int order = -1;
                string? format = null;
                foreach (var named in colAttr.NamedArguments)
                {
                    if (named.Key == "Order" && named.Value.Value is int o) order = o;
                    else if (named.Key == "Format" && named.Value.Value is string f) format = f;
                    else if (named.Key == "ConverterType" && named.Value.Value is INamedTypeSymbol convType)
                    {
                        converterTypeName = convType.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
                    }
                }
                column = new ColumnMapping(headerName, order, format);
            }

            var underlying = member.Type;
            var isNullable = false;
            if (underlying is INamedTypeSymbol nts && nts.IsGenericType
                && nts.ConstructedFrom.SpecialType == SpecialType.System_Nullable_T)
            {
                isNullable = true;
                underlying = nts.TypeArguments[0];
            }
            else if (member.NullableAnnotation == NullableAnnotation.Annotated
                     && underlying.IsReferenceType)
            {
                isNullable = true;
            }

            // v0.3.x: nullable value types are not yet generator-supported
            // (no Set(int, T?) overloads). They trip NXLS0006 honestly
            // until the next slice adds nullable-aware emit.
            // A custom converter (decision I-58) overrides the built-in
            // type check — the converter is responsible for read/write.
            var supported = converterTypeName is not null
                || (!isNullable && IsSupportedPropertyType(underlying));

            var locFrom = member.Locations.FirstOrDefault();
            var loc = LocationFrom(locFrom) ?? new PropertyLocation("<unknown>", 0, 0, 0);

            var wp = new WorksheetProperty(
                name: member.Name,
                fullTypeName: member.Type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
                underlyingSpecialType: underlying.SpecialType,
                isNullable: isNullable,
                column: column,
                isIgnored: ignored,
                location: loc,
                typeIsSupported: supported,
                converterTypeFullName: converterTypeName);

            // An override / `new` shadow keeps the base column position but
            // adopts the most-derived declaration's metadata.
            if (indexByName.TryGetValue(member.Name, out int existing))
                result[existing] = wp;
            else
            {
                indexByName[member.Name] = result.Count;
                result.Add(wp);
            }
        }
        return result;
    }

    private static bool IsSupportedPropertyType(ITypeSymbol type)
    {
        // Special types — scalars known to the runtime.
        var bySpecial = type.SpecialType switch
        {
            SpecialType.System_String => true,
            SpecialType.System_Boolean => true,
            SpecialType.System_Byte => true,
            SpecialType.System_SByte => true,
            SpecialType.System_Int16 => true,
            SpecialType.System_UInt16 => true,
            SpecialType.System_Int32 => true,
            SpecialType.System_UInt32 => true,
            SpecialType.System_Int64 => true,
            SpecialType.System_UInt64 => true,
            SpecialType.System_Single => true,
            SpecialType.System_Double => true,
            SpecialType.System_Decimal => true,
            SpecialType.System_DateTime => true,
            _ => false,
        };
        if (bySpecial) return true;

        // Non-special types we model: DateOnly, TimeOnly, TimeSpan.
        // (Guid still falls through and trips NXLS0006 — pending
        // ICell.SetGuid in a future slice.)
        var fullName = type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
        return fullName is "global::System.DateOnly"
            or "global::System.TimeOnly"
            or "global::System.TimeSpan";
    }

    /// <summary>
    /// Returns the text of a <c>row.Set(col, record.Prop)</c>-style call
    /// for the given property. Casts the property to a type that
    /// unambiguously resolves to one of <see cref="IRow"/>'s
    /// <c>Set</c> overloads (string, bool, int, long, double, decimal).
    /// Returns <c>null</c> if the type is not supported (callers should
    /// gate on <see cref="WorksheetProperty.TypeIsSupported"/>).
    /// </summary>
    private static string? FormatSetCall(WorksheetProperty p, int columnIndex)
    {
        // Custom converter (decision I-58) overrides the built-in
        // type-to-Set-overload mapping. Cached as a static field on
        // the generated extension class — see EmitConverterFields.
        if (p.ConverterTypeFullName is not null)
        {
            return $"s_conv_{p.Name}.Write(row.Cell({columnIndex}), record.{p.Name});";
        }

        var castedExpression = p.UnderlyingSpecialType switch
        {
            SpecialType.System_String => $"record.{p.Name}",
            SpecialType.System_Boolean => $"record.{p.Name}",
            SpecialType.System_Int32 => $"record.{p.Name}",
            SpecialType.System_Int16
                or SpecialType.System_Byte
                or SpecialType.System_SByte
                or SpecialType.System_UInt16 => $"(int)record.{p.Name}",
            SpecialType.System_Int64 => $"record.{p.Name}",
            SpecialType.System_UInt32 => $"(long)record.{p.Name}",
            SpecialType.System_UInt64 => $"(long)record.{p.Name}",  // wraps for values > long.MaxValue
            SpecialType.System_Double => $"record.{p.Name}",
            SpecialType.System_Single => $"(double)record.{p.Name}",
            SpecialType.System_Decimal => $"record.{p.Name}",
            SpecialType.System_DateTime => $"record.{p.Name}",
            _ => null,
        };

        // Non-special types we model (DateOnly, TimeOnly, TimeSpan) — keyed
        // off the fully-qualified type name since they have no SpecialType.
        if (castedExpression is null)
        {
            castedExpression = p.FullTypeName switch
            {
                "global::System.DateOnly"
                    or "global::System.TimeOnly"
                    or "global::System.TimeSpan" => $"record.{p.Name}",
                _ => null,
            };
        }

        return castedExpression is null
            ? null
            : $"row.Set({columnIndex}, {castedExpression});";
    }

    /// <summary>
    /// Returns the text of a cell-read expression that converts the
    /// indicated property's column on <paramref name="rowVar"/> back to
    /// the property's declared type. Required (non-nullable) properties
    /// throw <c>WorkbookException</c> on missing / wrong-typed cells.
    /// Returns <c>null</c> if the property type is not supported.
    /// </summary>
    private static string? FormatReadExpression(WorksheetProperty p, string rowVar)
    {
        // cell access: row.Cell(col_PropName)
        var cellExpr = $"{rowVar}.Cell(col_{p.Name})";

        // Custom converter (decision I-58) overrides the built-in
        // cell-kind dispatch. Cached as a static field — see
        // EmitConverterFields.
        if (p.ConverterTypeFullName is not null)
        {
            return $"s_conv_{p.Name}.Read({cellExpr})";
        }

        // shared throw expression for "required cell missing / wrong type".
        // The header is embedded in a generated *interpolated* string literal, so it
        // must escape the backslash/quote (literal-breaking) AND the braces (would be
        // read as interpolation holes) — otherwise a [Column] header containing any of
        // ", \, { or } emits code that does not compile.
        string escapedHeader = p.Column!.HeaderName
            .Replace("\\", "\\\\").Replace("\"", "\\\"")
            .Replace("{", "{{").Replace("}", "}}");
        string ThrowExpr(string expected) =>
            $"throw new global::NetXlsx.WorkbookException($\"Row {{r}} column '{escapedHeader}' expected {expected} but got {{({cellExpr}).Kind}}.\")";

        return p.UnderlyingSpecialType switch
        {
            SpecialType.System_String => $"{cellExpr}.GetString()",
            SpecialType.System_Boolean => $"({cellExpr}.GetBool() ?? {ThrowExpr("bool")})",
            SpecialType.System_Int32 => $"(int)({cellExpr}.GetNumber() ?? {ThrowExpr("int")})",
            SpecialType.System_Int16 => $"(short)({cellExpr}.GetNumber() ?? {ThrowExpr("short")})",
            SpecialType.System_Byte => $"(byte)({cellExpr}.GetNumber() ?? {ThrowExpr("byte")})",
            SpecialType.System_SByte => $"(sbyte)({cellExpr}.GetNumber() ?? {ThrowExpr("sbyte")})",
            SpecialType.System_UInt16 => $"(ushort)({cellExpr}.GetNumber() ?? {ThrowExpr("ushort")})",
            SpecialType.System_Int64 => $"(long)({cellExpr}.GetNumber() ?? {ThrowExpr("long")})",
            SpecialType.System_UInt32 => $"(uint)({cellExpr}.GetNumber() ?? {ThrowExpr("uint")})",
            SpecialType.System_UInt64 => $"(ulong)({cellExpr}.GetNumber() ?? {ThrowExpr("ulong")})",
            SpecialType.System_Single => $"(float)({cellExpr}.GetNumber() ?? {ThrowExpr("float")})",
            SpecialType.System_Double => $"({cellExpr}.GetNumber() ?? {ThrowExpr("double")})",
            SpecialType.System_Decimal => $"(decimal)({cellExpr}.GetNumber() ?? {ThrowExpr("decimal")})",
            SpecialType.System_DateTime => $"({cellExpr}.GetDate() ?? {ThrowExpr("DateTime")})",
            _ => p.FullTypeName switch
            {
                "global::System.DateOnly" => $"({cellExpr}.GetDateOnly() ?? {ThrowExpr("DateOnly")})",
                "global::System.TimeOnly" => $"({cellExpr}.GetTime() ?? {ThrowExpr("TimeOnly")})",
                "global::System.TimeSpan" => $"({cellExpr}.GetDuration() ?? {ThrowExpr("TimeSpan")})",
                _ => null,
            },
        };
    }

    private static bool HasDesignatedConstructor(INamedTypeSymbol type, TypeDeclarationSyntax syntax)
    {
        // Record primary constructor counts.
        if (syntax is RecordDeclarationSyntax rd && rd.ParameterList is not null)
            return true;

        // Public parameterless ctor (explicit or compiler-supplied on a class
        // with no explicit ctors).
        var explicitCtors = type.InstanceConstructors.Where(c => !c.IsImplicitlyDeclared).ToList();
        if (explicitCtors.Count == 0)
        {
            // Compiler-supplied default ctor for a class. Structs always
            // have implicit default; records without primary ctor too.
            return true;
        }
        return explicitCtors.Any(c => c.Parameters.Length == 0 && c.DeclaredAccessibility == Accessibility.Public);
    }

    private static TypeKindLite ClassifyKind(INamedTypeSymbol type)
    {
        if (type.IsRecord && type.IsValueType) return TypeKindLite.RecordStruct;
        if (type.IsRecord) return TypeKindLite.Record;
        if (type.IsValueType) return TypeKindLite.Struct;
        return TypeKindLite.Class;
    }

    private static string GetVisibility(GeneratorAttributeSyntaxContext ctx)
    {
        foreach (var attr in ctx.Attributes)
        {
            foreach (var named in attr.NamedArguments)
            {
                if (named.Key == "Visibility" && named.Value.Value is int v && v == 1)
                    return "public";
            }
        }
        return "internal";
    }

    private static PropertyLocation? LocationFrom(Location? loc)
    {
        if (loc is null || !loc.IsInSource) return null;
        var lineSpan = loc.GetLineSpan();
        return new PropertyLocation(
            lineSpan.Path,
            lineSpan.StartLinePosition.Line,
            lineSpan.StartLinePosition.Character,
            loc.SourceSpan.Length);
    }

    // ------------------------------------------------------------------
    // Property-level diagnostics (post-transform)
    // ------------------------------------------------------------------

    private static void EmitPropertyDiagnostics(SourceProductionContext spc, WorksheetModel model)
    {
        foreach (var p in model.Properties)
        {
            if (p.IsIgnored) continue;
            if (p.Column is null)
            {
                spc.ReportDiagnostic(Diagnostic.Create(
                    Diagnostics.UnmappedProperty,
                    p.Location.ToRoslynLocation(),
                    model.TypeName, p.Name));
                continue;
            }
            if (!p.TypeIsSupported)
            {
                spc.ReportDiagnostic(Diagnostic.Create(
                    Diagnostics.UnsupportedPropertyType,
                    p.Location.ToRoslynLocation(),
                    model.TypeName, p.Name, p.FullTypeName));
            }
            if (p.Column.Format is not null && !PassesFormatStringSmokeCheck(p.Column.Format, out var reason))
            {
                spc.ReportDiagnostic(Diagnostic.Create(
                    Diagnostics.MalformedFormatString,
                    p.Location.ToRoslynLocation(),
                    model.TypeName, p.Name, p.Column.Format, reason));
            }
        }
    }

    /// <summary>
    /// Conservative structural smoke check on an Excel number-format string.
    /// Catches: empty input, control characters, unbalanced <c>[...]</c>
    /// regions. Does <em>not</em> validate the full Excel format-string
    /// grammar — that's Excel's responsibility at render time. A
    /// "literal-only" rejection was considered and rejected: Excel's
    /// date/time format letters (<c>h</c>, <c>m</c>, <c>d</c>, <c>y</c>,
    /// <c>s</c>, etc.) collide with ordinary English text often enough
    /// that the check would have a misleading false-positive rate. The
    /// diagnostic name and message describe exactly what is and is not
    /// checked.
    /// </summary>
    private static bool PassesFormatStringSmokeCheck(string fmt, out string reason)
    {
        if (string.IsNullOrEmpty(fmt))
        {
            reason = "format string is empty";
            return false;
        }
        int bracketDepth = 0;
        foreach (var ch in fmt)
        {
            if (ch == '\0' || (ch < 0x20 && ch != '\t'))
            {
                reason = string.Format(System.Globalization.CultureInfo.InvariantCulture,
                    "contains control character U+{0:X4}", (int)ch);
                return false;
            }
            if (ch == '[') bracketDepth++;
            else if (ch == ']') bracketDepth--;
            if (bracketDepth < 0) { reason = "unmatched ']'"; return false; }
        }
        if (bracketDepth != 0)
        {
            reason = string.Format(System.Globalization.CultureInfo.InvariantCulture,
                "unmatched '[' (depth {0})", bracketDepth);
            return false;
        }
        reason = string.Empty;
        return true;
    }

    private static bool HasFatalEarlyDiagnostic(WorksheetModel m) =>
        !m.IsPartial || !m.HasDesignatedConstructor;

    private static Diagnostic BuildDiagnostic(DiagnosticInfo info)
    {
        var args = info.MessageArgs.Array.Cast<object?>().ToArray();
        var loc = info.Location?.ToRoslynLocation() ?? Location.None;
        return Diagnostic.Create(info.Descriptor, loc, args);
    }

    // ------------------------------------------------------------------
    // Code emission
    // ------------------------------------------------------------------

    private static string EmitExtensionsSource(WorksheetModel m)
    {
        var sb = new StringBuilder();
        sb.AppendLine("// <auto-generated by NetXlsx.SourceGen.WorksheetGenerator />");
        sb.AppendLine("#nullable enable");
        sb.AppendLine("#pragma warning disable CS1591 // missing XML doc on emitted code");
        sb.AppendLine();
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Collections.Generic;");
        sb.AppendLine("using NetXlsx;");
        sb.AppendLine();

        var hasNs = !string.IsNullOrEmpty(m.Namespace);
        if (hasNs)
        {
            sb.Append("namespace ").Append(m.Namespace).AppendLine(";");
            sb.AppendLine();
        }

        var extClassName = $"{m.TypeName}_SheetExtensions";
        sb.Append(m.Visibility).Append(" static class ").Append(extClassName).AppendLine();
        sb.AppendLine("{");

        // Cached converter instances per property (decision I-58).
        // One static readonly field per property with a configured
        // ConverterType. Allocates once at class-init; subsequent
        // AddRow / ReadRows calls reuse the instance.
        foreach (var p in m.Properties)
        {
            if (p.IsIgnored) continue;
            if (p.Column is null) continue;
            if (p.ConverterTypeFullName is null) continue;
            sb.Append("    private static readonly global::NetXlsx.ICellConverter<")
              .Append(p.FullTypeName)
              .Append("> s_conv_")
              .Append(p.Name)
              .Append(" = new ")
              .Append(p.ConverterTypeFullName)
              .AppendLine("();");
        }
        sb.AppendLine();

        // Properties that get a Set call, in source-declaration order
        // among mapped non-ignored properties. v0.3.x does NOT honor
        // [Column(Order)] for write ordering — the attribute is parsed
        // and validated (NXLS0001 catches duplicates) but the write
        // sequence is currently declaration-order. Reorder-on-write
        // lands with the styling slice when format-string emission
        // arrives.
        var writableProps = new List<WorksheetProperty>();
        foreach (var p in m.Properties)
        {
            if (p.IsIgnored) continue;
            if (p.Column is null) continue;     // NXLS0004 (warning) already fired
            if (!p.TypeIsSupported) continue;   // NXLS0006 already fired
            writableProps.Add(p);
        }

        // ---- AddRow: real body, no [Obsolete] -----------------------------
        sb.AppendLine("    /// <summary>");
        sb.Append("    /// Appends one <see cref=\"").Append(m.FullyQualifiedName).AppendLine("\"/> as a row of the supplied sheet.");
        sb.AppendLine("    /// </summary>");
        sb.Append("    public static void AddRow(this global::NetXlsx.ISheet sheet, ").Append(m.FullyQualifiedName).AppendLine(" record)");
        sb.AppendLine("    {");
        sb.AppendLine("        if (sheet is null) throw new global::System.ArgumentNullException(nameof(sheet));");
        sb.AppendLine("        if (record is null) throw new global::System.ArgumentNullException(nameof(record));");
        sb.AppendLine("        var row = sheet.AppendRow();");
        for (int i = 0; i < writableProps.Count; i++)
        {
            var callText = FormatSetCall(writableProps[i], i + 1);
            // FormatSetCall returned non-null because TypeIsSupported gated us.
            sb.Append("        ").AppendLine(callText);
        }
        sb.AppendLine("    }");
        sb.AppendLine();

        // ---- AddRows: foreach -> AddRow, no [Obsolete] --------------------
        sb.AppendLine("    /// <summary>");
        sb.Append("    /// Appends a sequence of <see cref=\"").Append(m.FullyQualifiedName).AppendLine("\"/> records as rows.");
        sb.AppendLine("    /// </summary>");
        sb.Append("    public static void AddRows(this global::NetXlsx.ISheet sheet, global::System.Collections.Generic.IEnumerable<").Append(m.FullyQualifiedName).AppendLine("> records)");
        sb.AppendLine("    {");
        sb.AppendLine("        if (sheet is null) throw new global::System.ArgumentNullException(nameof(sheet));");
        sb.AppendLine("        if (records is null) throw new global::System.ArgumentNullException(nameof(records));");
        sb.AppendLine("        foreach (var record in records) { AddRow(sheet, record); }");
        sb.AppendLine("    }");
        sb.AppendLine();

        // ---- ReadRows: real body --------------------------------------
        sb.AppendLine("    /// <summary>");
        sb.Append("    /// Reads rows from the sheet as a sequence of <see cref=\"").Append(m.FullyQualifiedName).AppendLine("\"/>.");
        sb.AppendLine("    /// </summary>");
        sb.AppendLine("    /// <param name=\"sheet\">The sheet to read from.</param>");
        sb.AppendLine("    /// <param name=\"headerRow\">1-based header row (default <c>1</c>). Header-less reading is deferred to a future milestone — pass null and a <c>NotSupportedException</c> is thrown.</param>");
        sb.Append("    public static global::System.Collections.Generic.IEnumerable<").Append(m.FullyQualifiedName).Append("> ReadRows(this global::NetXlsx.ISheet sheet, int? headerRow = 1)").AppendLine();
        sb.AppendLine("    {");
        sb.AppendLine("        if (sheet is null) throw new global::System.ArgumentNullException(nameof(sheet));");
        sb.AppendLine("        if (!headerRow.HasValue)");
        sb.AppendLine("            throw new global::System.NotSupportedException(\"Header-less ReadRows is deferred to v2 (decision I-46). Supply a header row.\");");
        sb.AppendLine();
        sb.AppendLine("        var headerRowIndex = headerRow.Value;");
        sb.AppendLine("        var headerRowObj = sheet.Row(headerRowIndex);");
        sb.AppendLine("        var headerColumns = new global::System.Collections.Generic.Dictionary<string, int>(global::System.StringComparer.OrdinalIgnoreCase);");
        sb.AppendLine("        for (int hc = 1; hc <= global::NetXlsx.CellAddress.MaxColumn; hc++)");
        sb.AppendLine("        {");
        sb.AppendLine("            var headerCell = headerRowObj.Cell(hc);");
        sb.AppendLine("            if (headerCell.Kind == global::NetXlsx.CellKind.Empty) break;");
        sb.AppendLine("            headerColumns[headerCell.GetString()] = hc;");
        sb.AppendLine("        }");
        sb.AppendLine();
        // Resolve each mapped property's column via [Column(Name)] -> header lookup.
        foreach (var p in writableProps)
        {
            var headerName = p.Column!.HeaderName.Replace("\\", "\\\\").Replace("\"", "\\\"");
            sb.Append("        if (!headerColumns.TryGetValue(\"").Append(headerName).Append("\", out int col_").Append(p.Name).AppendLine("))");
            sb.Append("            throw new global::NetXlsx.WorkbookException(\"Header '").Append(headerName).AppendLine("' not found in header row.\");");
        }
        sb.AppendLine();
        sb.AppendLine("        var lastRow = sheet.LastRowNumber;   // 1-based; 0 when the sheet is empty (decision I-85)");
        sb.AppendLine("        for (int r = headerRowIndex + 1; r <= lastRow; r++)");
        sb.AppendLine("        {");
        sb.AppendLine("            var row = sheet.Row(r);");
        // End-of-data heuristic: row is fully empty for the mapped columns.
        sb.AppendLine("            bool anyMappedCellHasValue = false;");
        foreach (var p in writableProps)
        {
            sb.Append("            if (row.Cell(col_").Append(p.Name).AppendLine(").Kind != global::NetXlsx.CellKind.Empty) anyMappedCellHasValue = true;");
        }
        sb.AppendLine("            if (!anyMappedCellHasValue) continue;");
        sb.AppendLine();
        sb.Append("            yield return new ").Append(m.FullyQualifiedName).AppendLine();
        sb.AppendLine("            {");
        foreach (var p in writableProps)
        {
            sb.Append("                ").Append(p.Name).Append(" = ").Append(FormatReadExpression(p, "row")).AppendLine(",");
        }
        sb.AppendLine("            };");
        sb.AppendLine("        }");
        sb.AppendLine("    }");

        sb.AppendLine("}");
        return sb.ToString();
    }

    private static string Sanitize(string fullyQualifiedName)
    {
        var sb = new StringBuilder(fullyQualifiedName.Length);
        foreach (var ch in fullyQualifiedName)
        {
            if (ch == ':' || ch == '<' || ch == '>' || ch == ',' || ch == ' ') sb.Append('_');
            else sb.Append(ch);
        }
        return sb.ToString();
    }
}
