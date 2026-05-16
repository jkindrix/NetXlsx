// Verifies the generator emits the expected extension class shape.

using System.Linq;
using FluentAssertions;
using Xunit;

namespace NetXlsx.Tests.SourceGen;

public class EmissionTests
{
    [Fact]
    public void Emits_Extensions_Class_For_Valid_Worksheet()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial class Row
{
    [Column(""A"")] public string A { get; set; } = """";
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratedSources.Should().NotBeEmpty();

        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().Contain("internal static class Row_SheetExtensions");
        generated.Should().Contain("public static void AddRow(");
        generated.Should().Contain("public static void AddRows(");
        generated.Should().Contain("public static global::System.Collections.Generic.IEnumerable<global::T.Row> ReadRows(");
        generated.Should().Contain("int? headerRow = 1");
    }

    [Fact]
    public void Visibility_Public_Opts_Generated_Class_Public()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet(Visibility = WorksheetVisibility.Public)]
public partial class Row
{
    [Column(""A"")] public string A { get; set; } = """";
}";
        var output = GeneratorHarness.Run(src);
        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().Contain("public static class Row_SheetExtensions");
    }

    [Fact]
    public void Fatal_Diagnostics_Skip_Emission()
    {
        // Type not partial → NXLS0005 → no generated file.
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public class Row
{
    [Column(""A"")] public string A { get; set; } = """";
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().Contain(d => d.Id == "NXLS0005");
        output.GeneratedSources.Should().BeEmpty(
            "fatal type-level diagnostics short-circuit code emission to avoid a downstream compile error cascade");
    }

    [Fact]
    public void Cross_Assembly_Worksheet_Types_Are_Ignored()
    {
        // No [Worksheet] in this compilation — generator emits nothing.
        // Mirrors the I5 contract: only the current compilation is scanned.
        const string src = @"
using NetXlsx;
namespace T;
public class NotAnnotated
{
    public string A { get; set; } = """";
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratedSources.Should().BeEmpty();
        output.GeneratorDiagnostics.Should().BeEmpty();
    }
}
