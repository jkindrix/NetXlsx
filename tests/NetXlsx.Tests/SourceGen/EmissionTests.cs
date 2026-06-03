// Verifies the generator emits the expected extension class shape.

using System.Linq;
using AwesomeAssertions;
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
    public void Inherited_Public_Columns_Are_Mapped_Base_First()
    {
        // A [Worksheet] type that inherits mappable public properties from
        // a base class must include them (previously they were silently
        // dropped). Base columns lead the derived type's own.
        const string src = @"
using NetXlsx;
namespace T;
public abstract class Base
{
    [Column(""Id"")] public int Id { get; set; }
}
[Worksheet]
public partial class Row : Base
{
    [Column(""Name"")] public string Name { get; set; } = """";
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().NotContain(d => d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Error);

        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().Contain("record.Id");
        generated.Should().Contain("record.Name");
        // Base property leads (column 1) before the derived one (column 2).
        generated.IndexOf("record.Id", System.StringComparison.Ordinal)
            .Should().BeLessThan(generated.IndexOf("record.Name", System.StringComparison.Ordinal));
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
    public void Emitted_AddRow_AddRows_Are_Not_Obsolete()
    {
        // v0.3.x: AddRow / AddRows now have real bodies that call
        // ISheet.AppendRow + IRow.Set. They must NOT be decorated
        // [Obsolete] (CS0619) — consumers can call them.
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

        // AddRow body uses AppendRow().Set(...) form.
        generated.Should().Contain("sheet.AppendRow()");
        generated.Should().Contain("row.Set(1, record.A);");

        // The [Obsolete] decoration only applies to ReadRows now.
        var addRowIdx = generated.IndexOf("public static void AddRow(", System.StringComparison.Ordinal);
        var beforeAddRow = generated.Substring(System.Math.Max(0, addRowIdx - 200), System.Math.Min(200, addRowIdx));
        beforeAddRow.Should().NotContain("[global::System.Obsolete",
            "AddRow must not be decorated [Obsolete] in v0.3.x — its body is real");
    }

    [Fact]
    public void Emitted_ReadRows_Has_Real_Body_No_Longer_Obsolete()
    {
        // v0.5.x: ReadRows landed; the [Obsolete(error: true)] decoration
        // is gone, and the body resolves headers + yields records.
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

        generated.Should().Contain("ReadRows(");
        // No [Obsolete] anywhere in the file now.
        generated.Should().NotContain("[global::System.Obsolete(",
            "all emitted methods now have real bodies");
        // ReadRows uses the header-resolution shape we designed.
        generated.Should().Contain("headerColumns.TryGetValue");
        generated.Should().Contain("yield return new global::T.Row");
    }

    [Fact]
    public void Calling_ReadRows_No_Longer_Produces_CS0619()
    {
        const string src = @"
using System.Collections.Generic;
using System.Linq;
using NetXlsx;
namespace T;

[Worksheet]
public partial class Row
{
    [Column(""A"")] public string A { get; set; } = """";
}

public static class Consumer
{
    public static IEnumerable<Row> Use(ISheet sheet) => Row_SheetExtensions.ReadRows(sheet);
}";
        var output = GeneratorHarness.Run(src);
        output.CompilationDiagnostics.Should().NotContain(d => d.Id == "CS0619",
            "ReadRows is now a real method, not an obsolete stub");
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

    [Fact]
    public void Header_With_Quote_Backslash_And_Braces_Emits_Compilable_Code()
    {
        // Regression: the wrong-type throw expression interpolates the [Column] header
        // into a generated *interpolated* string literal. A header containing ", \, {
        // or } must be escaped (quote/backslash break the literal; braces would be read
        // as interpolation holes) or the generated code fails to compile. Header here is
        // a"b\c{d}.
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial class Row
{
    [Column(""a\""b\\c{d}"")] public int A { get; set; }
}";
        var output = GeneratorHarness.Run(src);

        output.GeneratedSources.Should().NotBeEmpty();
        // Ignore the harness's one benign gap — the generated ReadRows body reaches
        // sheet.Underlying (NPOI XSSFSheet), which this minimal reference set omits
        // (CS0012). Any OTHER error means the header escaping is wrong: an unescaped
        // quote/backslash is a syntax error, an unescaped brace is CS0103 (the literal
        // is read as an interpolation hole referencing an undefined name).
        output.CompilationDiagnostics.Should().NotContain(
            d => d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Error && d.Id != "CS0012",
            "a header with quote, backslash, or brace characters must be escaped into the generated literal");

        // The escaped header is present verbatim in the generated throw expression.
        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().Contain(@"a\""b\\c{{d}}");
    }
}
