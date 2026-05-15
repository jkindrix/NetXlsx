using System;
using BenchmarkDotNet.Running;

namespace NetXlsx.Benchmarks;

/// <summary>
/// Entry point. Two modes:
/// <list type="bullet">
/// <item><c>spike-N</c> — run pre-impl spike N (1, 2, or 3). Spike 4 lives in
///       <c>spikes/NetXlsx.AotSpike/</c> because it requires its own publish.</item>
/// <item>anything else (or no args) — BenchmarkDotNet switcher across all <c>[*Bench]</c> classes.</item>
/// </list>
/// </summary>
public static class Program
{
    public static int Main(string[] args)
    {
        if (args.Length > 0)
        {
            switch (args[0])
            {
                case "spike-1": return Spike1_StyleDedupFeasibility.Run();
                case "spike-2": return Spike2_StreamingBackPressure.Run();
                case "spike-3": return Spike3_AsyncWrappingCost.Run();
                case "spike-all":
                    var r = Spike1_StyleDedupFeasibility.Run();
                    if (r != 0) return r;
                    r = Spike2_StreamingBackPressure.Run();
                    if (r != 0) return r;
                    return Spike3_AsyncWrappingCost.Run();
            }
        }
        BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
        return 0;
    }
}
