// R-4 regression suite: positional records construct through the record
// primary constructor in the generated ReadRows (previously the emit was
// object-initializer-only — CS7036 in the .g.cs with no diagnostic), and
// genuinely unconstructible shapes fire NXLS0007 instead of emitting
// non-compiling code. Includes the compile-AND-RUN round-trip the
// remediation ledger requires.

using System.Collections;
using System.Linq;
using AwesomeAssertions;
using Microsoft.CodeAnalysis;
using Xunit;

namespace NetXlsx.Tests.SourceGen;

public class PositionalRecordEmissionTests
{
    // The README's flagship typed-mapping shape (README.md "Typed
    // export / import"), minus the Format argument — byte-for-byte the
    // R-4 repro.
    private const string PositionalRecordSource = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial record SalesRow(
    [property: Column(""Region"")]  string Region,
    [property: Column(""Revenue"")] decimal Revenue);";

    [Fact]
    public void Positional_Record_Emits_Primary_Ctor_Construction_That_Compiles()
    {
        var output = GeneratorHarness.RunWithFullReferences(PositionalRecordSource, out _);

        output.GeneratorDiagnostics.Should().BeEmpty(
            "a positional record is a fully supported shape — no diagnostic may fire");
        output.CompilationDiagnostics.Should().NotContain(
            d => d.Severity == DiagnosticSeverity.Error,
            "the generated ReadRows must construct via the primary ctor, not an object initializer (CS7036)");

        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().Contain("yield return new global::T.SalesRow(");
        generated.Should().Contain("Region: ");
        generated.Should().Contain("Revenue: ");
    }

    [Fact]
    public void Positional_Record_RoundTrips_AddRows_To_ReadRows_At_Runtime()
    {
        var output = GeneratorHarness.RunWithFullReferences(PositionalRecordSource, out var compilation);
        output.CompilationDiagnostics.Should().NotContain(d => d.Severity == DiagnosticSeverity.Error);

        var asm = GeneratorHarness.EmitAndLoad(compilation);
        var rowType = asm.GetType("T.SalesRow")!;
        var ext = asm.GetType("T.SalesRow_SheetExtensions")!;

        var records = System.Array.CreateInstance(rowType, 2);
        records.SetValue(System.Activator.CreateInstance(rowType, "West", 1234.56m), 0);
        records.SetValue(System.Activator.CreateInstance(rowType, "East", 99m), 1);

        using var ms = new System.IO.MemoryStream();
        using (var wb = Workbook.Create())
        {
            var sheet = wb.AddSheet("Sales");
            var header = sheet.AppendRow();
            header.Set(1, "Region");
            header.Set(2, "Revenue");
            ext.GetMethod("AddRows")!.Invoke(null, new object[] { sheet, records });
            wb.Save(ms);
        }

        ms.Position = 0;
        using (var wb = Workbook.Open(ms))
        {
            var rows = (IEnumerable)ext.GetMethod("ReadRows")!.Invoke(null, new object?[] { wb["Sales"], 1 })!;
            var back = rows.Cast<object>().ToList();

            back.Should().HaveCount(2);
            back[0].Should().Be(records.GetValue(0),
                "records have value equality — the round-trip must reproduce them exactly");
            back[1].Should().Be(records.GetValue(1));
        }
    }

    [Fact]
    public void Positional_Record_With_Extra_Body_Property_Uses_Initializer_Suffix()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial record Row(
    [property: Column(""A"")] string A)
{
    [Column(""B"")] public int B { get; set; }
}";
        var output = GeneratorHarness.RunWithFullReferences(src, out _);

        output.GeneratorDiagnostics.Should().BeEmpty();
        output.CompilationDiagnostics.Should().NotContain(d => d.Severity == DiagnosticSeverity.Error);

        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().Contain("yield return new global::T.Row(");
        generated.Should().Contain("A: ");
        generated.Should().Contain("B = ");
    }

    [Fact]
    public void Positional_Readonly_Record_Struct_Compiles()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public readonly partial record struct Row(
    [property: Column(""A"")] string A,
    [property: Column(""N"")] double N);";
        var output = GeneratorHarness.RunWithFullReferences(src, out _);

        output.GeneratorDiagnostics.Should().BeEmpty();
        output.CompilationDiagnostics.Should().NotContain(d => d.Severity == DiagnosticSeverity.Error);

        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().Contain("yield return new global::T.Row(");
    }

    [Fact]
    public void Defaulted_Unmapped_Ctor_Param_Is_Omitted_And_Compiles()
    {
        // `Tag` is unmapped (NXLS0004 warning) but has a default value, so
        // ReadRows omits it from the ctor call — the type stays usable.
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial record Row(
    [property: Column(""A"")] string A,
    int Tag = 7);";
        var output = GeneratorHarness.RunWithFullReferences(src, out _);

        output.GeneratorDiagnostics.Should().NotContain(d => d.Id == "NXLS0007");
        output.GeneratorDiagnostics.Should().Contain(d => d.Id == "NXLS0004",
            "the unmapped Tag property still gets the explicit-mapping warning");
        output.CompilationDiagnostics.Should().NotContain(d => d.Severity == DiagnosticSeverity.Error);

        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().NotContain("Tag:", "a defaulted unmapped ctor parameter is omitted");
    }

    [Fact]
    public void Unmapped_Ctor_Param_Without_Default_Fires_NXLS0007_And_Skips_Emission()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial record Row(
    [property: Column(""A"")] string A,
    [property: Ignore] int NoDefault);";
        var output = GeneratorHarness.Run(src);

        output.GeneratorDiagnostics.Should().Contain(d =>
            d.Id == "NXLS0007" && d.GetMessage(null).Contains("NoDefault"));
        output.GeneratedSources.Should().BeEmpty(
            "an unconstructible type must not emit code that cannot compile (the pre-fix R-4 failure mode)");
    }

    [Fact]
    public void GetOnly_Mapped_Property_Fires_NXLS0007()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial class Row
{
    [Column(""A"")] public string A { get; } = """";
}";
        var output = GeneratorHarness.Run(src);

        output.GeneratorDiagnostics.Should().Contain(d =>
            d.Id == "NXLS0007" && d.GetMessage(null).Contains("'A'"));
        output.GeneratedSources.Should().BeEmpty();
    }

    [Fact]
    public void Required_Unmapped_Property_Fires_NXLS0007()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial class Row
{
    [Column(""A"")] public string A { get; set; } = """";
    [Ignore] public required int Mandatory { get; set; }
}";
        var output = GeneratorHarness.Run(src);

        output.GeneratorDiagnostics.Should().Contain(d =>
            d.Id == "NXLS0007" && d.GetMessage(null).Contains("Mandatory"));
        output.GeneratedSources.Should().BeEmpty(
            "an object initializer that does not set a required member is CS9035");
    }

    [Fact]
    public void Required_Mapped_Settable_Property_Compiles()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial class Row
{
    [Column(""A"")] public required string A { get; set; }
}";
        var output = GeneratorHarness.RunWithFullReferences(src, out _);

        output.GeneratorDiagnostics.Should().BeEmpty();
        output.CompilationDiagnostics.Should().NotContain(d => d.Severity == DiagnosticSeverity.Error,
            "a required property the initializer covers is constructible");
    }

    [Fact]
    public void Parameterless_Record_Still_Uses_Object_Initializer()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial record Row
{
    [Column(""A"")] public string A { get; set; } = """";
}";
        var output = GeneratorHarness.RunWithFullReferences(src, out _);

        output.GeneratorDiagnostics.Should().BeEmpty();
        output.CompilationDiagnostics.Should().NotContain(d => d.Severity == DiagnosticSeverity.Error);

        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().NotContain("yield return new global::T.Row(");
        generated.Should().Contain("A = ");
    }
}
