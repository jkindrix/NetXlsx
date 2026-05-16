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
        var result = new List<WorksheetProperty>();
        foreach (var member in type.GetMembers().OfType<IPropertySymbol>())
        {
            ct.ThrowIfCancellationRequested();
            if (member.DeclaredAccessibility != Accessibility.Public) continue;
            if (member.IsStatic) continue;
            if (member.IsIndexer) continue;

            var attrs = member.GetAttributes();
            var ignored = attrs.Any(a => a.AttributeClass?.ToDisplayString() == IgnoreAttributeFullName);
            var colAttr = attrs.FirstOrDefault(a => a.AttributeClass?.ToDisplayString() == ColumnAttributeFullName);

            ColumnMapping? column = null;
            if (colAttr is not null && colAttr.ConstructorArguments.Length > 0)
            {
                var headerName = colAttr.ConstructorArguments[0].Value as string ?? member.Name;
                int order = -1;
                string? format = null;
                foreach (var named in colAttr.NamedArguments)
                {
                    if (named.Key == "Order" && named.Value.Value is int o) order = o;
                    else if (named.Key == "Format" && named.Value.Value is string f) format = f;
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

            var supported = IsSupportedPropertyType(underlying);

            var locFrom = member.Locations.FirstOrDefault();
            var loc = LocationFrom(locFrom) ?? new PropertyLocation("<unknown>", 0, 0, 0);

            result.Add(new WorksheetProperty(
                name: member.Name,
                fullTypeName: member.Type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
                isNullable: isNullable,
                column: column,
                isIgnored: ignored,
                location: loc,
                typeIsSupported: supported));
        }
        return result;
    }

    private static bool IsSupportedPropertyType(ITypeSymbol type)
    {
        return type.SpecialType switch
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
            _ => type.ToDisplayString() switch
            {
                "System.DateOnly" => true,
                "System.TimeOnly" => true,
                "System.TimeSpan" => true,
                "System.Guid" => true,
                _ => false,
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

        // The emitted methods are decorated with [Obsolete(error: true)]
        // until ISheet gains its members in milestone 2. The point: any
        // consumer who takes a preview dependency and tries to call these
        // methods gets a *compile error* (CS0619) with the message below,
        // not a runtime NotImplementedException. The throwing body remains
        // as a belt-and-suspenders backstop for reflection-based call
        // paths that bypass the [Obsolete] check.
        const string StubMessage =
            "NetXlsx typed-mapping extension methods are emitted but not yet executable: ISheet implementation lands in milestone 2 (see docs/design.md §6.4 and CHANGELOG.md). Track the milestone before relying on this API.";

        const string ObsoleteAttr =
            "[global::System.Obsolete(\"" + StubMessage + "\", error: true)]";

        sb.AppendLine("    /// <summary>");
        sb.Append("    /// Writes one <see cref=\"").Append(m.FullyQualifiedName).AppendLine("\"/> as a row of the supplied sheet.");
        sb.AppendLine("    /// </summary>");
        sb.Append("    ").AppendLine(ObsoleteAttr);
        sb.Append("    public static void AddRow(this global::NetXlsx.ISheet sheet, ").Append(m.FullyQualifiedName).AppendLine(" record)");
        sb.AppendLine("    {");
        sb.Append("        throw new global::System.NotImplementedException(\"").Append(StubMessage).AppendLine("\");");
        sb.AppendLine("    }");
        sb.AppendLine();

        sb.AppendLine("    /// <summary>");
        sb.Append("    /// Writes a sequence of <see cref=\"").Append(m.FullyQualifiedName).AppendLine("\"/> records as appended rows.");
        sb.AppendLine("    /// </summary>");
        sb.Append("    ").AppendLine(ObsoleteAttr);
        sb.Append("    public static void AddRows(this global::NetXlsx.ISheet sheet, global::System.Collections.Generic.IEnumerable<").Append(m.FullyQualifiedName).AppendLine("> records)");
        sb.AppendLine("    {");
        sb.Append("        throw new global::System.NotImplementedException(\"").Append(StubMessage).AppendLine("\");");
        sb.AppendLine("    }");
        sb.AppendLine();

        sb.AppendLine("    /// <summary>");
        sb.Append("    /// Reads rows from the sheet as a sequence of <see cref=\"").Append(m.FullyQualifiedName).AppendLine("\"/>.");
        sb.AppendLine("    /// </summary>");
        sb.AppendLine("    /// <param name=\"sheet\">The sheet to read from.</param>");
        sb.AppendLine("    /// <param name=\"headerRow\">1-based header row (default <c>1</c>). Pass <c>null</c> for header-less mode.</param>");
        sb.Append("    ").AppendLine(ObsoleteAttr);
        sb.Append("    public static global::System.Collections.Generic.IEnumerable<").Append(m.FullyQualifiedName).Append("> ReadRows(this global::NetXlsx.ISheet sheet, int? headerRow = 1)").AppendLine();
        sb.AppendLine("    {");
        sb.Append("        throw new global::System.NotImplementedException(\"").Append(StubMessage).AppendLine("\");");
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
