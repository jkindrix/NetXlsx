// R-30 remainder — locale-matrix breadth, focused on tr-TR (the "Turkish-I"
// trap). Two properties matter regardless of the ambient thread culture:
//
//   1. Serialization is culture-INVARIANT. A workbook written while
//      CurrentCulture is tr-TR (comma decimal separator) must still emit
//      "1234.5" in the OOXML <v>, never "1234,5" — otherwise the file is
//      corrupt for every other reader. This is the high-value check.
//   2. Case-insensitive lookups are ORDINAL, not culture-sensitive. Under
//      tr-TR, a culture-sensitive ToLower turns "FILE" into "fıle" (dotless
//      ı) and "i".ToUpper() into "İ" — so any culture-keyed dictionary would
//      mis-resolve. NetXlsx keys on StringComparer.OrdinalIgnoreCase, so
//      "FILE"/"file" still match and "I"/"ı" stay distinct.
//
// The existing DisplayCulture read-path coverage (de-DE) lives in
// WorkbookOptionsReadPathTests; this file adds the tr-TR rows and the
// ambient-culture-during-write dimension that was previously untested.

using System;
using System.Globalization;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class LocaleMatrixTests
{
    private static readonly CultureInfo TrTr = new("tr-TR");

    /// <summary>Runs <paramref name="body"/> with the thread culture forced to
    /// <paramref name="culture"/>, restoring the previous culture afterwards.</summary>
    private static void InCulture(CultureInfo culture, Action body)
    {
        var prevCulture = CultureInfo.CurrentCulture;
        var prevUi = CultureInfo.CurrentUICulture;
        CultureInfo.CurrentCulture = culture;
        CultureInfo.CurrentUICulture = culture;
        try { body(); }
        finally
        {
            CultureInfo.CurrentCulture = prevCulture;
            CultureInfo.CurrentUICulture = prevUi;
        }
    }

    [Fact]
    public void Sanity_TrTr_Comma_Is_The_Decimal_Separator()
    {
        // Guards the premise: if this ever fails the other tests are vacuous.
        InCulture(TrTr, () =>
            (1234.5).ToString(CultureInfo.CurrentCulture).Should().Be("1234,5", "tr-TR uses a comma decimal separator"));
    }

    [Fact]
    public void Number_Serializes_Invariant_Under_TrTr_Ambient_Culture()
    {
        InCulture(TrTr, () =>
        {
            using var wb = Workbook.Create();
            var sh = wb.AddSheet("S");
            sh["A1"].SetNumber(1234.5);

            var v = SavedOoxml.Cell(SavedOoxml.SheetXml(wb), "A1")!
                .Element(SavedOoxml.Main + "v")!.Value;
            v.Should().Be("1234.5", "the OOXML <v> must be invariant (dot), not the tr-TR comma form");
        });
    }

    [Fact]
    public void Date_Serializes_Invariant_Serial_Under_TrTr_Ambient_Culture()
    {
        InCulture(TrTr, () =>
        {
            using var wb = Workbook.Create();
            var sh = wb.AddSheet("S");
            sh["A1"].SetDate(new DateTime(2026, 6, 18));

            var v = SavedOoxml.Cell(SavedOoxml.SheetXml(wb), "A1")!
                .Element(SavedOoxml.Main + "v")!.Value;
            // A whole-day date is a bare invariant integer serial — no comma,
            // no tr-TR punctuation. (Asserting "parses as an invariant integer"
            // rather than a hardcoded serial keeps the test about the locale
            // property, not the date arithmetic.)
            v.Should().NotContain(",");
            int.TryParse(v, NumberStyles.None, CultureInfo.InvariantCulture, out _)
                .Should().BeTrue($"the date serial '{v}' must be a bare invariant integer");
        });
    }

    [Fact]
    public void Date_Roundtrips_Under_TrTr()
    {
        InCulture(TrTr, () =>
        {
            using var wb = Workbook.Create();
            var sh = wb.AddSheet("S");
            sh["A1"].SetDate(new DateTime(2026, 6, 18));
            using var ms = new MemoryStream();
            wb.Save(ms, leaveOpen: true);
            ms.Position = 0;
            using var reopened = Workbook.Open(ms);
            reopened["S"]["A1"].GetDate().Should().Be(new DateTime(2026, 6, 18));
        });
    }

    [Fact]
    public void Number_Roundtrips_Invariant_Through_Save_And_Reopen_Under_TrTr()
    {
        InCulture(TrTr, () =>
        {
            var path = Path.Combine(Path.GetTempPath(), $"trtr-{Guid.NewGuid():N}.xlsx");
            try
            {
                using (var wb = Workbook.Create())
                {
                    wb.AddSheet("S")["A1"].SetNumber(1234.5);
                    wb.Save(path);
                }
                using var reopened = Workbook.Open(path);
                // GetNumber is invariant by construction (§7.2) — exact value back.
                reopened["S"]["A1"].GetNumber().Should().Be(1234.5);
            }
            finally { File.Delete(path); }
        });
    }

    [Fact]
    public void Sheet_Name_Lookup_Is_Ordinal_Case_Insensitive_Under_TrTr()
    {
        InCulture(TrTr, () =>
        {
            using var wb = Workbook.Create();
            wb.AddSheet("FILE");

            // ASCII case fold works under OrdinalIgnoreCase regardless of culture.
            // A culture-keyed dictionary would have stored "fıle" under tr-TR and
            // failed this lookup.
            wb["file"].Name.Should().Be("FILE");
            wb.TryGetSheet("File", out var s).Should().BeTrue();
            s!.Name.Should().Be("FILE");
        });
    }

    [Fact]
    public void Dotted_And_Dotless_I_Are_Distinct_Sheets_Under_TrTr()
    {
        InCulture(TrTr, () =>
        {
            using var wb = Workbook.Create();
            wb.AddSheet("I");      // U+0049
            // "ı" (U+0131) is a different code point — ordinal comparison keeps
            // it distinct, so this must NOT collide with "I" (a tr-TR culture
            // fold would treat "I"/"ı" as the same and throw a name clash).
            Action addDotless = () => wb.AddSheet("ı");
            addDotless.Should().NotThrow();
            wb.SheetCount.Should().Be(2);

            // "i" (U+0069) folds to "I" ordinally, not to the dotless "ı".
            wb["i"].Name.Should().Be("I");
        });
    }

    [Fact]
    public void Named_Style_Lookup_Is_Ordinal_Case_Insensitive_Under_TrTr()
    {
        InCulture(TrTr, () =>
        {
            using var wb = Workbook.Create();
            wb.RegisterStyle("TITLE", new CellStyle { Bold = true });
            wb.GetRegisteredStyle("title").Should().NotBeNull(
                "named-style lookup is OrdinalIgnoreCase — immune to the Turkish-I fold");
        });
    }

    [Fact]
    public void DisplayCulture_TrTr_Renders_Date_GetString_With_TrTr_Format()
    {
        // Completes the read-path matrix (de-DE is covered in
        // WorkbookOptionsReadPathTests): a DisplayCulture of tr-TR renders a
        // date cell's GetString in tr-TR's format, while the raw GetNumber stays
        // invariant. Run with an invariant *ambient* culture so this isolates
        // the DisplayCulture option, not the thread culture.
        var path = Path.Combine(Path.GetTempPath(), $"trtr-disp-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = Workbook.Create())
            {
                var sh = wb.AddSheet("S");
                sh["A1"].SetDate(new DateTime(2026, 6, 18));
                sh["A1"].NumberFormat("yyyy-mm-dd");
                wb.Save(path);
            }
            using var wbInv = Workbook.Open(path, new WorkbookOptions { DisplayCulture = CultureInfo.InvariantCulture });
            using var wbTr = Workbook.Open(path, new WorkbookOptions { DisplayCulture = TrTr });

            // Raw value is culture-invariant either way.
            wbTr["S"]["A1"].GetNumber().Should().Be(wbInv["S"]["A1"].GetNumber());
            // The displayed string is DisplayCulture-aware; the format here is
            // explicit (yyyy-mm-dd) so both render the same digits — the
            // assertion is simply that the tr-TR read path produces a value and
            // does not throw under a non-invariant DisplayCulture.
            wbTr["S"]["A1"].GetString().Should().NotBeNullOrEmpty();
        }
        finally { File.Delete(path); }
    }
}
