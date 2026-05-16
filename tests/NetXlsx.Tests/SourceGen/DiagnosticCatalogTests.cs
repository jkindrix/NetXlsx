// One test per diagnostic ID per design §6.12 / decision I5.
// Each verifies the diagnostic fires on its trigger case.

using System.Linq;
using FluentAssertions;
using Xunit;

namespace NetXlsx.Tests.SourceGen;

public class DiagnosticCatalogTests
{
    [Fact]
    public void NXLS0001_Duplicate_Column_Order_Fires()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial class Row
{
    [Column(""A"", Order = 0)] public string A { get; set; } = """";
    [Column(""B"", Order = 0)] public string B { get; set; } = """";
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().ContainSingle(d => d.Id == "NXLS0001");
    }

    [Fact]
    public void NXLS0002_Missing_Designated_Constructor_Fires()
    {
        // Class with only a parameterized public ctor — no parameterless one.
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial class Row
{
    public Row(int x) { }
    [Column(""A"")] public string A { get; set; } = """";
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().Contain(d => d.Id == "NXLS0002");
    }

    [Fact]
    public void NXLS0002_Record_Primary_Constructor_Does_Not_Fire()
    {
        // Per I5 / §6.12: record primary ctor satisfies the rule.
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial record Row([property: Column(""A"")] string A);";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().NotContain(d => d.Id == "NXLS0002");
    }

    [Fact]
    public void NXLS0003_Unbalanced_Bracket_Fires()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial class Row
{
    [Column(""A"", Format = ""[[unbalanced"")] public double A { get; set; }
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().Contain(d => d.Id == "NXLS0003");
    }

    [Fact]
    public void NXLS0003_Embedded_Control_Char_Fires()
    {
        // Per the design comment in PassesFormatStringSmokeCheck: a
        // "literal-only" rejection was considered and rejected because
        // Excel's date/time letters (h, m, d, y, s) collide with ordinary
        // English too often. Control characters remain a clean trigger.
        const string src = "using NetXlsx;\nnamespace T;\n[Worksheet]\npublic partial class Row\n{\n    [Column(\"A\", Format = \"0.00\\u0001\")] public double A { get; set; }\n}";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().Contain(d => d.Id == "NXLS0003");
    }

    [Fact]
    public void NXLS0003_Valid_Excel_Format_Does_Not_Fire()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial class Row
{
    [Column(""A"", Format = ""$#,##0.00;[Red]-$#,##0.00"")] public decimal A { get; set; }
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().NotContain(d => d.Id == "NXLS0003");
    }

    [Fact]
    public void NXLS0004_Unmapped_Property_Warns()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial class Row
{
    public string Unmapped { get; set; } = """";
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().Contain(d =>
            d.Id == "NXLS0004"
            && d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Warning);
    }

    [Fact]
    public void NXLS0005_Type_Not_Partial_Fires()
    {
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
    }

    [Fact]
    public void NXLS0006_Unsupported_Property_Type_Fires()
    {
        const string src = @"
using NetXlsx;
namespace T;
public class CustomType { }

[Worksheet]
public partial class Row
{
    [Column(""A"")] public CustomType A { get; set; } = new();
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().Contain(d => d.Id == "NXLS0006");
    }

    [Fact]
    public void Clean_Worksheet_Type_Produces_No_Diagnostics()
    {
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial class Row
{
    [Column(""Region"", Order = 0)] public string Region { get; set; } = """";
    [Column(""Revenue"", Format = ""$#,##0.00"")] public decimal Revenue { get; set; }
    [Ignore] public string InternalNotes { get; set; } = """";
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().BeEmpty();
    }
}
