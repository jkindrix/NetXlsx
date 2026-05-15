using BenchmarkDotNet.Running;

namespace NetXlsx.Benchmarks;

/// <summary>
/// Entry point for BenchmarkDotNet. Real benchmarks (pre-impl spikes #1-4
/// plus per-PR regression benches per design §5) land during v1.0
/// implementation. Run with: <c>dotnet run -c Release -- --filter '*'</c>.
/// </summary>
public static class Program
{
    public static void Main(string[] args)
    {
        BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
    }
}
