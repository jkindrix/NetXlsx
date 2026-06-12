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
    private static readonly MetadataReference[] s_fullReferences = BuildFullReferences();
    private static int s_assemblyCounter;

    public static GeneratorOutput Run(string userSource) =>
        Run(userSource, s_baseReferences, out _);

    /// <summary>
    /// Like <see cref="Run(string)"/> but compiles against the test
    /// process's full reference closure, so the generated code must
    /// compile with ZERO errors (no CS0012 carve-outs) and the resulting
    /// compilation can be emitted and executed via
    /// <see cref="EmitAndLoad"/>.
    /// </summary>
    public static GeneratorOutput RunWithFullReferences(string userSource, out CSharpCompilation updatedCompilation) =>
        Run(userSource, s_fullReferences, out updatedCompilation);

    private static GeneratorOutput Run(string userSource, MetadataReference[] references, out CSharpCompilation updatedCompilation)
    {
        var syntaxTree = CSharpSyntaxTree.ParseText(userSource);
        var compilation = CSharpCompilation.Create(
            assemblyName: $"TestAssembly{System.Threading.Interlocked.Increment(ref s_assemblyCounter)}",
            syntaxTrees: new[] { syntaxTree },
            references: references,
            options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, nullableContextOptions: NullableContextOptions.Enable));

        var generator = new WorksheetGenerator();
        var driver = (GeneratorDriver)CSharpGeneratorDriver.Create(generator);
        driver = driver.RunGeneratorsAndUpdateCompilation(compilation, out var updated, out var genDiags);
        updatedCompilation = (CSharpCompilation)updated;

        var compDiags = updated.GetDiagnostics();
        var generated = driver.GetRunResult().Results
            .SelectMany(r => r.GeneratedSources)
            .Select(s => (s.HintName, s.SourceText.ToString()))
            .ToImmutableArray();

        return new GeneratorOutput(genDiags, compDiags, generated);
    }

    /// <summary>
    /// Emits the compilation to memory and loads it into the test
    /// process (default load context — assembly references such as
    /// NetXlsx resolve to the already-loaded copies). Throws with the
    /// emit diagnostics when the compilation does not build.
    /// </summary>
    public static Assembly EmitAndLoad(CSharpCompilation compilation)
    {
        using var ms = new System.IO.MemoryStream();
        var result = compilation.Emit(ms);
        if (!result.Success)
        {
            var errors = string.Join("\n", result.Diagnostics
                .Where(d => d.Severity == DiagnosticSeverity.Error)
                .Select(d => d.ToString()));
            throw new System.InvalidOperationException($"Emit failed:\n{errors}");
        }
        return Assembly.Load(ms.ToArray());
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

    /// <summary>
    /// The full reference closure of the running test process: every
    /// trusted-platform assembly plus NetXlsx and its references
    /// (force-loaded so DocumentFormat.OpenXml etc. are present even if
    /// the test run has not touched them yet).
    /// </summary>
    private static MetadataReference[] BuildFullReferences()
    {
        var paths = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);

        if (System.AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") is string tpa)
        {
            foreach (var p in tpa.Split(System.IO.Path.PathSeparator))
                if (p.Length > 0) paths.Add(p);
        }

        var netXlsx = typeof(NetXlsx.WorksheetAttribute).Assembly;
        paths.Add(netXlsx.Location);
        foreach (var name in netXlsx.GetReferencedAssemblies())
        {
            try
            {
                var dep = Assembly.Load(name);
                if (!dep.IsDynamic && dep.Location.Length > 0) paths.Add(dep.Location);
            }
            catch (System.IO.FileNotFoundException)
            {
                // A reference the runtime can't resolve standalone — the
                // TPA list already carries the framework facades we need.
            }
        }

        return paths.Select(p => (MetadataReference)MetadataReference.CreateFromFile(p)).ToArray();
    }
}
