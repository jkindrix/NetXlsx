// Minimal Roslyn harness for testing the WorksheetGenerator.
// Compiles a user-source snippet against the NetXlsx assembly +
// the source generator, captures emitted diagnostics + generated files.

using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Reflection;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using NetXlsx.SourceGen;

namespace NetXlsx.Tests.SourceGen;

internal sealed record GeneratorOutput(
    ImmutableArray<Diagnostic> GeneratorDiagnostics,
    ImmutableArray<Diagnostic> CompilationDiagnostics,
    ImmutableArray<(string HintName, string Source)> GeneratedSources);

internal static class GeneratorHarness
{
    private static readonly MetadataReference[] s_baseReferences = BuildBaseReferences();

    public static GeneratorOutput Run(string userSource)
    {
        var syntaxTree = CSharpSyntaxTree.ParseText(userSource);
        var compilation = CSharpCompilation.Create(
            assemblyName: "TestAssembly",
            syntaxTrees: new[] { syntaxTree },
            references: s_baseReferences,
            options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, nullableContextOptions: NullableContextOptions.Enable));

        var generator = new WorksheetGenerator();
        var driver = (GeneratorDriver)CSharpGeneratorDriver.Create(generator);
        driver = driver.RunGeneratorsAndUpdateCompilation(compilation, out var updated, out var genDiags);

        var compDiags = updated.GetDiagnostics();
        var generated = driver.GetRunResult().Results
            .SelectMany(r => r.GeneratedSources)
            .Select(s => (s.HintName, s.SourceText.ToString()))
            .ToImmutableArray();

        return new GeneratorOutput(genDiags, compDiags, generated);
    }

    private static MetadataReference[] BuildBaseReferences()
    {
        var list = new List<MetadataReference>
        {
            MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(System.Linq.Enumerable).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(System.Collections.Generic.IEnumerable<>).Assembly.Location),
            // NetXlsx — gives the test access to [Worksheet], [Column], [Ignore], ISheet
            MetadataReference.CreateFromFile(typeof(NetXlsx.WorksheetAttribute).Assembly.Location),
        };

        // System.Runtime — required for primitives to resolve under net8+
        var runtimePath = System.IO.Path.GetDirectoryName(typeof(object).Assembly.Location);
        if (runtimePath is not null)
        {
            foreach (var dll in new[] { "System.Runtime.dll", "netstandard.dll", "System.Collections.dll" })
            {
                var full = System.IO.Path.Combine(runtimePath, dll);
                if (System.IO.File.Exists(full))
                    list.Add(MetadataReference.CreateFromFile(full));
            }
        }

        return list.ToArray();
    }
}
