// Workbook.SuggestSheetName — design line 160.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class SuggestSheetNameTests
{
    [Fact]
    public void Returns_Proposed_Verbatim_When_No_Collision()
    {
        using var wb = Workbook.Create();
        Workbook.SuggestSheetName(wb, "Sales").Should().Be("Sales");
    }

    [Fact]
    public void Appends_Numeric_Suffix_On_Collision()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Sales");
        Workbook.SuggestSheetName(wb, "Sales").Should().Be("Sales (2)");
    }

    [Fact]
    public void Walks_Suffixes_Until_Unused()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Sales");
        wb.AddSheet("Sales (2)");
        wb.AddSheet("Sales (3)");
        Workbook.SuggestSheetName(wb, "Sales").Should().Be("Sales (4)");
    }

    [Fact]
    public void Is_Case_Insensitive_For_Collision_Check()
    {
        using var wb = Workbook.Create();
        wb.AddSheet("Sales");
        // "SALES" collides with "Sales" (Excel sheet-name lookup is
        // case-insensitive — matches AddSheet's duplicate detection).
        Workbook.SuggestSheetName(wb, "SALES").Should().Be("SALES (2)");
    }

    [Fact]
    public void Truncates_To_31_Chars_Preserving_Suffix()
    {
        using var wb = Workbook.Create();
        var thirtyOne = new string('x', 31);
        wb.AddSheet(thirtyOne);

        var result = Workbook.SuggestSheetName(wb, thirtyOne);
        result.Length.Should().BeLessThanOrEqualTo(31);
        result.Should().EndWith(" (2)");
    }

    [Fact]
    public void Sanitizes_Invalid_Characters()
    {
        using var wb = Workbook.Create();
        Workbook.SuggestSheetName(wb, "Bad/Name").Should().Be("Bad_Name");
    }

    [Fact]
    public void Falls_Back_To_GUID_Suffix_When_All_Numeric_Suffixes_Collide()
    {
        // The numeric-suffix loop runs 2..9999. After 9,998 collisions
        // the fallback path appends a "_<8-hex>" tag. Filling 9,998
        // real sheets is slow; use a stub IWorkbook that says "yes,
        // every name you propose already exists." This exercises the
        // defensive path the reviewer flagged as untested at the
        // boundary.
        var stub = new AlwaysCollidingWorkbook();
        var result = Workbook.SuggestSheetName(stub, "Sales");

        // GUID fallback shape: base + "_" + 8 hex chars.
        result.Length.Should().BeLessThanOrEqualTo(31);
        result.Should().StartWith("Sales_");
        var tag = result.Substring("Sales_".Length);
        tag.Length.Should().Be(8);
        tag.Should().MatchRegex("^[0-9a-fA-F]{8}$",
            "the fallback uses Guid.NewGuid().ToString(\"N\").Substring(0, 8)");
    }

    [Fact]
    public void GUID_Fallback_Truncates_Long_Bases_To_Fit_31_Chars()
    {
        var stub = new AlwaysCollidingWorkbook();
        var longName = new string('a', 31);   // already at the cap

        var result = Workbook.SuggestSheetName(stub, longName);

        result.Length.Should().BeLessThanOrEqualTo(31,
            "the GUID-fallback path truncates the base to MaxSheetNameLength - 9 " +
            "so the appended '_<8-hex>' suffix still fits");
        result.Should().EndWith(string.Concat("_", result.AsSpan(result.Length - 8)));
    }

    [Fact]
    public void Throws_On_Null_Inputs()
    {
        using var wb = Workbook.Create();
        Action a = () => Workbook.SuggestSheetName(null!, "x");
        Action b = () => Workbook.SuggestSheetName(wb, null!);
        a.Should().Throw<ArgumentNullException>();
        b.Should().Throw<ArgumentNullException>();
    }
}

/// <summary>
/// IWorkbook stub that reports every <see cref="TryGetSheet"/> call
/// as a hit. Used by the SuggestSheetName GUID-fallback test to
/// force the numeric-suffix loop to exhaust deterministically
/// without allocating 9,998 real sheets.
/// </summary>
internal sealed class AlwaysCollidingWorkbook : IWorkbook
{
    public int SheetCount => 0;
    public ISheet this[string name] => throw new NotImplementedException();
    public ISheet this[int index] => throw new NotImplementedException();
    public ISheet AddSheet(string name) => throw new NotImplementedException();

    public bool TryGetSheet(string name, [MaybeNullWhen(false)] out ISheet sheet)
    {
        // Always-collide: the stub claims every queried name is taken,
        // which is exactly the worst-case the SuggestSheetName loop
        // is meant to survive.
        sheet = null!;
        return true;
    }

    public void Save(Stream stream, bool leaveOpen = true) => throw new NotImplementedException();
    public void Save(string path) => throw new NotImplementedException();
    public Task SaveAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default) => throw new NotImplementedException();
    public Task SaveAsync(string path, CancellationToken ct = default) => throw new NotImplementedException();

    public INamedRange AddNamedRange(string name, string formula, string? sheetScope = null) => throw new NotImplementedException();
    public IReadOnlyList<INamedRange> NamedRanges => throw new NotImplementedException();

    public void Protect(WorkbookProtection? options = null) => throw new NotImplementedException();
    public void ProtectWithPassword(string password, WorkbookProtection? options = null) => throw new NotImplementedException();
    public void Unprotect() => throw new NotImplementedException();
    public bool IsProtected => false;

    public void RegisterStyle(string name, CellStyle style) => throw new NotImplementedException();
    public CellStyle? GetRegisteredStyle(string name) => null;
    public IReadOnlyCollection<string> RegisteredStyleNames => Array.Empty<string>();
    public StylePoolDiagnostics GetStylePoolDiagnostics() => default;

    public NPOI.XSSF.UserModel.XSSFWorkbook Underlying => throw new NotImplementedException();

    public void Dispose() { }
}
