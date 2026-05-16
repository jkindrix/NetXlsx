using System;
using System.IO;
using System.Threading.Tasks;
using NetXlsx.Cookbook.Recipes;

namespace NetXlsx.Cookbook;

/// <summary>
/// Cookbook entry point. Recipes are also invoked from
/// <c>tests/NetXlsx.GoldenFiles/</c> — each recipe is a static class
/// with a <c>Run(string outputPath)</c> method.
/// </summary>
public static class Program
{
    private static readonly (string Name, Func<string, Task> Run)[] s_recipes =
    {
        ("hello-workbook", HelloWorkbook.Run),
        ("tabular-export", TabularExport.Run),
        ("typed-export", TypedExport.Run),
        ("time-and-duration", TimeAndDuration.Run),
        ("styled-report", StyledReport.Run),
        ("cell-errors", CellErrors.Run),
    };

    public static async Task<int> Main(string[] args)
    {
        if (args.Length == 0 || args[0] is "-h" or "--help")
        {
            PrintHelp();
            return args.Length == 0 ? 1 : 0;
        }

        var recipeName = args[0];
        var outputPath = args.Length > 1
            ? args[1]
            : Path.Combine(Path.GetTempPath(), $"{recipeName}.xlsx");

        foreach (var (name, run) in s_recipes)
        {
            if (name == recipeName)
            {
                await run(outputPath).ConfigureAwait(false);
                Console.WriteLine($"OK — wrote {outputPath}");
                return 0;
            }
        }

        Console.Error.WriteLine($"Unknown recipe: {recipeName}");
        PrintHelp();
        return 2;
    }

    private static void PrintHelp()
    {
        Console.WriteLine("Usage: cookbook <recipe> [output-path]");
        Console.WriteLine();
        Console.WriteLine("Recipes (v0.2.0):");
        foreach (var (name, _) in s_recipes)
        {
            Console.WriteLine($"  {name}");
        }
        Console.WriteLine();
        Console.WriteLine("The full v1.0 set is listed in docs/design.md §8.1.");
    }
}
