// Coverage for the v1.1 custom-converter slice (decision I-58).
// Verifies the source generator:
//   1. Recognizes ColumnAttribute.ConverterType and treats the property
//      as supported even when the declared type is not in the built-in
//      scalar set.
//   2. Emits a cached `static readonly ICellConverter<T> s_conv_X` field.
//   3. Routes the write call through `s_conv_X.Write(row.Cell(col), record.X)`.
//   4. Routes the read call through `s_conv_X.Read(row.Cell(col_X))`.
//   5. Does not fire NXLS0006 for the otherwise-unsupported type.

using System.Linq;
using AwesomeAssertions;
using NetXlsx.Tests.SourceGen;
using Xunit;

namespace NetXlsx.Tests;

public class CustomConverterTests
{
    [Fact]
    public void Generator_Emits_Cached_Converter_Field_For_Property_With_ConverterType()
    {
        const string src = @"
using System.Collections.Generic;
using NetXlsx;

namespace T;

public sealed class TagsConverter : ICellConverter<List<string>>
{
    public void Write(ICell cell, List<string> value) => cell.SetString(string.Join("";"", value));
    public List<string> Read(ICell cell) => new List<string>(cell.GetString().Split(';'));
}

[Worksheet]
public partial record Item
{
    [Column(""Tags"", ConverterType = typeof(TagsConverter))]
    public List<string> Tags { get; init; } = new();
}";
        var output = GeneratorHarness.Run(src);
        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().Contain(
            "private static readonly global::NetXlsx.ICellConverter<global::System.Collections.Generic.List<string>> s_conv_Tags = new global::T.TagsConverter();");
    }

    [Fact]
    public void Generator_Routes_Write_Through_Converter()
    {
        const string src = @"
using System.Collections.Generic;
using NetXlsx;
namespace T;
public sealed class TagsConverter : ICellConverter<List<string>>
{
    public void Write(ICell cell, List<string> value) { }
    public List<string> Read(ICell cell) => new();
}
[Worksheet]
public partial record Item
{
    [Column(""Tags"", ConverterType = typeof(TagsConverter))]
    public List<string> Tags { get; init; } = new();
}";
        var output = GeneratorHarness.Run(src);
        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().Contain("s_conv_Tags.Write(row.Cell(1), record.Tags);");
    }

    [Fact]
    public void Generator_Routes_Read_Through_Converter()
    {
        const string src = @"
using System.Collections.Generic;
using NetXlsx;
namespace T;
public sealed class TagsConverter : ICellConverter<List<string>>
{
    public void Write(ICell cell, List<string> value) { }
    public List<string> Read(ICell cell) => new();
}
[Worksheet]
public partial record Item
{
    [Column(""Tags"", ConverterType = typeof(TagsConverter))]
    public List<string> Tags { get; init; } = new();
}";
        var output = GeneratorHarness.Run(src);
        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        generated.Should().Contain("Tags = s_conv_Tags.Read(row.Cell(col_Tags))");
    }

    [Fact]
    public void Generator_Does_Not_Fire_NXLS0006_For_Property_With_Converter()
    {
        const string src = @"
using System.Collections.Generic;
using NetXlsx;
namespace T;
public sealed class TagsConverter : ICellConverter<List<string>>
{
    public void Write(ICell cell, List<string> value) { }
    public List<string> Read(ICell cell) => new();
}
[Worksheet]
public partial record Item
{
    // List<string> is not in the built-in supported set; without
    // ConverterType this would fire NXLS0006.
    [Column(""Tags"", ConverterType = typeof(TagsConverter))]
    public List<string> Tags { get; init; } = new();
}";
        var output = GeneratorHarness.Run(src);
        output.GeneratorDiagnostics.Should().NotContain(d => d.Id == "NXLS0006",
            "a configured ConverterType overrides the built-in type-support check");
    }

    [Fact]
    public void Generator_Mixed_Property_Set_Still_Works()
    {
        // A worksheet with both a built-in-supported property (string)
        // and a converter-property (List<string>). Verifies the converter
        // path doesn't disturb the normal Set/Read emission.
        const string src = @"
using System.Collections.Generic;
using NetXlsx;
namespace T;
public sealed class TagsConverter : ICellConverter<List<string>>
{
    public void Write(ICell cell, List<string> value) { }
    public List<string> Read(ICell cell) => new();
}
[Worksheet]
public partial record Item
{
    [Column(""Name"")] public string Name { get; init; } = """";
    [Column(""Tags"", ConverterType = typeof(TagsConverter))]
    public List<string> Tags { get; init; } = new();
}";
        var output = GeneratorHarness.Run(src);
        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        // Built-in property uses the normal Set call.
        generated.Should().Contain("row.Set(1, record.Name);");
        // Converter property uses the converter call.
        generated.Should().Contain("s_conv_Tags.Write(row.Cell(2), record.Tags);");
    }

    [Fact]
    public void Generator_Emit_Includes_Converter_Property_Even_When_Type_Has_No_SpecialType()
    {
        // List<string> has SpecialType.None; the previous emit pipeline
        // gated writableProps on TypeIsSupported. Verify the converter
        // unblocks emission rather than the property being silently skipped.
        const string src = @"
using System.Collections.Generic;
using NetXlsx;
namespace T;
public sealed class TagsConverter : ICellConverter<List<string>>
{
    public void Write(ICell cell, List<string> value) { }
    public List<string> Read(ICell cell) => new();
}
[Worksheet]
public partial record Item
{
    [Column(""Tags"", ConverterType = typeof(TagsConverter))]
    public List<string> Tags { get; init; } = new();
}";
        var output = GeneratorHarness.Run(src);
        var generated = string.Concat(output.GeneratedSources.Select(s => s.Source));
        // If the property had been skipped, AddRow would emit nothing
        // between the AppendRow and the closing brace. Verify the
        // converter call shows up in the AddRow body.
        var addRowStart = generated.IndexOf("public static void AddRow(", System.StringComparison.Ordinal);
        var addRowEnd = generated.IndexOf("public static void AddRows(", System.StringComparison.Ordinal);
        addRowStart.Should().BeGreaterThan(-1);
        addRowEnd.Should().BeGreaterThan(addRowStart);
        var addRowBody = generated.Substring(addRowStart, addRowEnd - addRowStart);
        addRowBody.Should().Contain("s_conv_Tags.Write");
    }
}
