// Coverage for the v1.1 named-styles slice: IWorkbook.RegisterStyle /
// GetRegisteredStyle / RegisteredStyleNames; ICell.ApplyNamedStyle;
// IRange.ApplyNamedStyle.

using System;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class NamedStyleTests
{
    private static readonly string[] s_abNames = new[] { "A", "B" };
    private static readonly string[] s_threeNames = new[] { "Header", "Body", "Accent" };


    [Fact]
    public void RegisteredStyleNames_Is_Empty_By_Default()
    {
        using var wb = Workbook.Create();
        wb.RegisteredStyleNames.Should().BeEmpty();
    }

    [Fact]
    public void Register_Then_Get_Returns_Same_Style()
    {
        using var wb = Workbook.Create();
        var s = new CellStyle { Bold = true, FontSize = 14 };
        wb.RegisterStyle("Header", s);
        wb.GetRegisteredStyle("Header").Should().Be(s);
    }

    [Fact]
    public void GetRegisteredStyle_Is_Case_Insensitive()
    {
        using var wb = Workbook.Create();
        wb.RegisterStyle("Header", new CellStyle { Bold = true });
        wb.GetRegisteredStyle("header").Should().NotBeNull();
        wb.GetRegisteredStyle("HEADER").Should().NotBeNull();
    }

    [Fact]
    public void GetRegisteredStyle_Returns_Null_For_Unknown_Name()
    {
        using var wb = Workbook.Create();
        wb.GetRegisteredStyle("Missing").Should().BeNull();
    }

    [Fact]
    public void RegisterStyle_Replaces_Existing()
    {
        using var wb = Workbook.Create();
        wb.RegisterStyle("X", new CellStyle { Bold = true });
        wb.RegisterStyle("X", new CellStyle { Italic = true });
        var s = wb.GetRegisteredStyle("X");
        s!.Bold.Should().BeNull();
        s.Italic.Should().Be(true);
    }

    [Fact]
    public void RegisteredStyleNames_Lists_All_Registered()
    {
        using var wb = Workbook.Create();
        wb.RegisterStyle("A", CellStyle.Default);
        wb.RegisterStyle("B", CellStyle.Default);
        wb.RegisteredStyleNames.Should().BeEquivalentTo(s_abNames);
    }

    [Fact]
    public void RegisterStyle_Rejects_Null_Name()
    {
        using var wb = Workbook.Create();
        Action act = () => wb.RegisterStyle(null!, CellStyle.Default);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void RegisterStyle_Rejects_Empty_Name()
    {
        using var wb = Workbook.Create();
        Action act = () => wb.RegisterStyle("", CellStyle.Default);
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void RegisterStyle_Rejects_Null_Style()
    {
        using var wb = Workbook.Create();
        Action act = () => wb.RegisterStyle("X", null!);
        act.Should().Throw<ArgumentNullException>();
    }

    // ---- Apply --------------------------------------------------------

    [Fact]
    public void Cell_ApplyNamedStyle_Sets_The_Cell_Style()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        wb.RegisterStyle("Header", new CellStyle { Bold = true, FontSize = 14 });
        sh["A1"].SetString("hello");
        sh["A1"].ApplyNamedStyle("Header");
        var got = sh["A1"].GetStyle();
        got.Bold.Should().Be(true);
        got.FontSize.Should().Be(14);
    }

    [Fact]
    public void Cell_ApplyNamedStyle_Throws_For_Unknown_Name()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh["A1"].ApplyNamedStyle("Missing");
        act.Should().Throw<ArgumentException>().WithMessage("*Missing*");
    }

    [Fact]
    public void Range_ApplyNamedStyle_Sets_Every_Cell()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        wb.RegisterStyle("Body", new CellStyle { Italic = true });
        sh.Range("A1:B2").ApplyNamedStyle("Body");
        sh["A1"].GetStyle().Italic.Should().Be(true);
        sh["B2"].GetStyle().Italic.Should().Be(true);
    }

    [Fact]
    public void Range_ApplyNamedStyle_Throws_For_Unknown_Name()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        Action act = () => sh.Range("A1:B2").ApplyNamedStyle("Missing");
        act.Should().Throw<ArgumentException>();
    }

    // ---- OOXML round-trip (v1.3 / I-67) -------------------------------

    [Fact]
    public void Registered_Names_Survive_File_Roundtrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"named-rt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.RegisterStyle("Header", new CellStyle { Bold = true, FontSize = 14 });
                wb.RegisterStyle("Body", new CellStyle { Italic = true });
                wb.RegisterStyle("Accent", new CellStyle { Bold = true, FontColor = Color.Red });
                wb.AddSheet("S");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                wb.RegisteredStyleNames.Should().BeEquivalentTo(s_threeNames);
                var header = wb.GetRegisteredStyle("Header");
                header.Should().NotBeNull();
                header!.Bold.Should().Be(true);
                header.FontSize.Should().Be(14);
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Roundtrip_Preserves_Style_Name_Case_Insensitive_Lookup()
    {
        var path = Path.Combine(Path.GetTempPath(), $"named-case-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.RegisterStyle("MyHeader", new CellStyle { Bold = true });
                wb.AddSheet("S");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                // Case-insensitive lookup preserved after rehydration.
                wb.GetRegisteredStyle("myheader").Should().NotBeNull();
                wb.GetRegisteredStyle("MYHEADER").Should().NotBeNull();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Workbook_Without_Named_Styles_Reads_Back_Empty()
    {
        var path = Path.Combine(Path.GetTempPath(), $"named-empty-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.AddSheet("S");
                wb["S"]["A1"].SetString("hello");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                // No named styles registered; the built-in "Normal"
                // entry that Excel creates is intentionally hidden
                // from the rehydration to avoid noisy "Normal" entries
                // in RegisteredStyleNames.
                wb.RegisteredStyleNames.Should().BeEmpty();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void RegisterStyle_Replacement_Updates_Existing_OOXML_Entry()
    {
        // RegisterStyle("X", ...) twice: second call overwrites both
        // the in-process map and the OOXML CT_CellStyle entry.
        var path = Path.Combine(Path.GetTempPath(), $"named-replace-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                wb.RegisterStyle("X", new CellStyle { Bold = true });
                wb.RegisterStyle("X", new CellStyle { Italic = true });
                wb.AddSheet("S");
                wb.Save(path);
            }
            using (var wb = Workbook.Open(path))
            {
                var s = wb.GetRegisteredStyle("X");
                s.Should().NotBeNull();
                // The second registration wins. Italic is preserved.
                s!.Italic.Should().Be(true);
                // OOXML should have exactly one entry named "X" — not
                // two from the duplicate registration.
                wb.RegisteredStyleNames.Should().ContainSingle();
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
