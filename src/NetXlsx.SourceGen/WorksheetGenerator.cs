using Microsoft.CodeAnalysis;

namespace NetXlsx.SourceGen;

/// <summary>
/// Scaffold placeholder for the <c>[Worksheet]</c> source generator
/// (design §6.9 / §6.12). The real generator — incremental, emitting
/// extension methods on <c>ISheet</c> / <c>IStreamingSheet</c> for typed
/// row mapping, with diagnostics <c>NXLS0001</c>–<c>NXLS0006</c> —
/// is implemented during v1.0.
/// </summary>
[Generator(LanguageNames.CSharp)]
public sealed class WorksheetGenerator : IIncrementalGenerator
{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        // Intentionally empty in scaffold. Real implementation lands in v1.0.
    }
}
