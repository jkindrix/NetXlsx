using System;
using System.Linq;
using FluentAssertions;
using Xunit;

namespace NetXlsx.PublicApi;

/// <summary>
/// Companion assertions to the <c>Microsoft.CodeAnalysis.PublicApiAnalyzers</c>
/// wiring on the main library. The analyzer catches additions at compile
/// time via <c>RS0016</c>/<c>RS0017</c> diagnostics (treated as errors —
/// see <c>.editorconfig</c>). These tests provide a runtime backstop:
/// a guarded enumeration of the public surface that fails if any public
/// type leaks before being deliberately added.
/// </summary>
public class PublicApiSnapshotTests
{
    [Fact]
    public void Library_Has_No_Public_Types_Yet()
    {
        // Scaffold: the library has no public types. This test pins that
        // baseline. When the first real public type lands, replace this
        // assertion with the full snapshot enumeration.
        var assembly = System.Reflection.Assembly.Load("NetXlsx");
        var publicTypes = assembly
            .GetExportedTypes()
            .Where(t => !IsCompilerGenerated(t))
            .ToList();

        publicTypes.Should().BeEmpty(
            "v1.0 public surface lands deliberately via the design §6 sketch; " +
            "until then, no public types should exist");
    }

    private static bool IsCompilerGenerated(Type t) =>
        t.GetCustomAttributes(typeof(System.Runtime.CompilerServices.CompilerGeneratedAttribute), false).Any();
}
