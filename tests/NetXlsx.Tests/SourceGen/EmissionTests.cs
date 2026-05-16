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
    public void Emitted_Methods_Are_Obsolete_Error_Until_Milestone_2()
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
        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().Contain("[global::System.Obsolete(", "consumer-level calls must fail at compile time, not runtime — see CHANGELOG.md");
        generated.Should().Contain("error: true",
            "the [Obsolete] decoration must be CS0619-level (error), not a soft warning");
    }

    [Fact]
    public void Calling_Emitted_Method_Produces_CS0619_Compile_Error()
    {
        // Verifies the contract end-to-end: a consumer that compiles
        // against the generator output gets a build break, not a
        // crash-at-runtime, when calling the obsolete-error stubs.
        const string src = @"
using NetXlsx;
namespace T;

[Worksheet]
public partial class Row
{
    [Column(""A"")] public string A { get; set; } = """";
}

public static class Consumer
{
    public static void Use(ISheet sheet)
    {
        Row_SheetExtensions.AddRow(sheet, new Row());
    }
}";
        var output = GeneratorHarness.Run(src);
        output.CompilationDiagnostics.Should().Contain(d => d.Id == "CS0619",
            "calling an [Obsolete(error: true)] method must produce CS0619 at compile time");
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
