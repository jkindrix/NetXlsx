// Comprehensive disposal matrix — every public mutating method and
// property accessor across IWorkbook / ISheet / IRow / ICell must throw
// ObjectDisposedException after the owning workbook is disposed
// (decision #42). Parameterized so the addition of a new public method
// is one MemberData row, not a copy-paste.

using System;
using System.Collections.Generic;
using System.IO;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class DisposedWorkbookMatrixTests
{
    // ---- IWorkbook ------------------------------------------------------

    public static IEnumerable<object[]> WorkbookOperations()
    {
        yield return new object[] { "SheetCount", (Action<IWorkbook>)(wb => { var _ = wb.SheetCount; }) };
        yield return new object[] { "this[string]", (Action<IWorkbook>)(wb => { var _ = wb["S"]; }) };
        yield return new object[] { "this[int]", (Action<IWorkbook>)(wb => { var _ = wb[0]; }) };
        yield return new object[] { "AddSheet", (Action<IWorkbook>)(wb => wb.AddSheet("Another")) };
        yield return new object[] { "TryGetSheet", (Action<IWorkbook>)(wb => wb.TryGetSheet("S", out _)) };
        yield return new object[] { "Save(stream)", (Action<IWorkbook>)(wb => wb.Save(new MemoryStream())) };
        yield return new object[] { "Save(path)", (Action<IWorkbook>)(wb => wb.Save(Path.Combine(Path.GetTempPath(), "x.xlsx"))) };
        yield return new object[] { "Underlying", (Action<IWorkbook>)(wb => { var _ = wb.Underlying; }) };
        yield return new object[] { "AddNamedRange", (Action<IWorkbook>)(wb => wb.AddNamedRange("X", "S!$A$1")) };
        yield return new object[] { "NamedRanges", (Action<IWorkbook>)(wb => { var _ = wb.NamedRanges; }) };
        yield return new object[] { "Protect", (Action<IWorkbook>)(wb => wb.Protect()) };
        yield return new object[] { "ProtectWithPassword", (Action<IWorkbook>)(wb => wb.ProtectWithPassword("p")) };
        yield return new object[] { "Unprotect", (Action<IWorkbook>)(wb => wb.Unprotect()) };
        yield return new object[] { "IsProtected", (Action<IWorkbook>)(wb => { var _ = wb.IsProtected; }) };
        yield return new object[] { "RegisterStyle", (Action<IWorkbook>)(wb => wb.RegisterStyle("H", CellStyle.Default)) };
        yield return new object[] { "GetRegisteredStyle", (Action<IWorkbook>)(wb => wb.GetRegisteredStyle("H")) };
        yield return new object[] { "RegisteredStyleNames", (Action<IWorkbook>)(wb => { var _ = wb.RegisteredStyleNames; }) };
        yield return new object[] { "GetStylePoolDiagnostics", (Action<IWorkbook>)(wb => wb.GetStylePoolDiagnostics()) };
    }

    [Theory]
    [MemberData(nameof(WorkbookOperations))]
    public void Workbook_Member_Throws_ObjectDisposed_After_Dispose(string memberName, Action<IWorkbook> op)
    {
        var wb = Workbook.Create();
        wb.AddSheet("S");
        wb.Dispose();

        Action act = () => op(wb);
        act.Should().Throw<ObjectDisposedException>(
            $"{memberName} must throw ObjectDisposedException after the workbook is disposed (decision #42)");
    }

    // ---- ISheet ---------------------------------------------------------

    public static IEnumerable<object[]> SheetOperations()
    {
        yield return new object[] { "Name", (Action<ISheet>)(s => { var _ = s.Name; }) };
        yield return new object[] { "Workbook", (Action<ISheet>)(s => { var _ = s.Workbook; }) };
        yield return new object[] { "this[A1]", (Action<ISheet>)(s => { var _ = s["A1"]; }) };
        yield return new object[] { "this[r,c]", (Action<ISheet>)(s => { var _ = s[1, 1]; }) };
        yield return new object[] { "AppendRow", (Action<ISheet>)(s => s.AppendRow()) };
        yield return new object[] { "Row(1)", (Action<ISheet>)(s => s.Row(1)) };
        yield return new object[] { "Underlying", (Action<ISheet>)(s => { var _ = s.Underlying; }) };
        yield return new object[] { "FreezeRows", (Action<ISheet>)(s => s.FreezeRows(1)) };
        yield return new object[] { "FreezeColumns", (Action<ISheet>)(s => s.FreezeColumns(1)) };
        yield return new object[] { "FreezePane", (Action<ISheet>)(s => s.FreezePane(1, 1)) };
        yield return new object[] { "MergeCells", (Action<ISheet>)(s => s.MergeCells("A1:B2")) };
        yield return new object[] { "UnmergeCells", (Action<ISheet>)(s => s.UnmergeCells("A1:B2")) };
        yield return new object[] { "MergedRanges", (Action<ISheet>)(s => { var _ = s.MergedRanges; }) };
        yield return new object[] { "Hidden get", (Action<ISheet>)(s => { var _ = s.Hidden; }) };
        yield return new object[] { "Hidden set", (Action<ISheet>)(s => { s.Hidden = true; }) };
        yield return new object[] { "ShowGridlines get", (Action<ISheet>)(s => { var _ = s.ShowGridlines; }) };
        yield return new object[] { "ShowGridlines set", (Action<ISheet>)(s => { s.ShowGridlines = false; }) };
        yield return new object[] { "Range(string)", (Action<ISheet>)(s => s.Range("A1:B2")) };
        yield return new object[] { "Range(r,c,r,c)", (Action<ISheet>)(s => s.Range(1, 1, 2, 2)) };
        yield return new object[] { "Column(int)", (Action<ISheet>)(s => s.Column(1)) };
        yield return new object[] { "Column(string)", (Action<ISheet>)(s => s.Column("A")) };
        yield return new object[] { "AddTable", (Action<ISheet>)(s => s.AddTable("A1:B2", "T")) };
        yield return new object[] { "Tables", (Action<ISheet>)(s => { var _ = s.Tables; }) };
        yield return new object[] { "TryGetTable", (Action<ISheet>)(s => s.TryGetTable("T", out _)) };
        yield return new object[] { "RemoveTable", (Action<ISheet>)(s => s.RemoveTable(new XssfTableStub())) };
        yield return new object[] { "AddPicture(3-arg)", (Action<ISheet>)(s => s.AddPicture("A1", new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, ImageFormat.Png)) };
        yield return new object[] { "AddPicture(2-arg)", (Action<ISheet>)(s => s.AddPicture("A1", new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A })) };
        yield return new object[] { "Protect", (Action<ISheet>)(s => s.Protect()) };
        yield return new object[] { "Unprotect", (Action<ISheet>)(s => s.Unprotect()) };
        yield return new object[] { "IsProtected", (Action<ISheet>)(s => { var _ = s.IsProtected; }) };
        yield return new object[] { "AddValidation", (Action<ISheet>)(s => s.AddValidation("A1:A5", DataValidation.IntegerBetween(1, 10))) };
        yield return new object[] { "SetAutoFilter", (Action<ISheet>)(s => s.SetAutoFilter("A1:B5")) };
        yield return new object[] { "ClearAutoFilter", (Action<ISheet>)(s => s.ClearAutoFilter()) };
        yield return new object[] { "HasAutoFilter", (Action<ISheet>)(s => { var _ = s.HasAutoFilter; }) };
        yield return new object[] { "AutoFilterRange", (Action<ISheet>)(s => { var _ = s.AutoFilterRange; }) };
    }

    [Theory]
    [MemberData(nameof(SheetOperations))]
    public void Sheet_Member_Throws_ObjectDisposed_After_Workbook_Dispose(string memberName, Action<ISheet> op)
    {
        var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        wb.Dispose();

        Action act = () => op(sheet);
        act.Should().Throw<ObjectDisposedException>(
            $"ISheet.{memberName} must throw ObjectDisposedException once its owning workbook is disposed");
    }

    // ---- IRow -----------------------------------------------------------

    public static IEnumerable<object[]> RowOperations()
    {
        yield return new object[] { "Index", (Action<IRow>)(r => { var _ = r.Index; }) };
        yield return new object[] { "Sheet", (Action<IRow>)(r => { var _ = r.Sheet; }) };
        yield return new object[] { "Cell(1)", (Action<IRow>)(r => r.Cell(1)) };
        yield return new object[] { "this[1]", (Action<IRow>)(r => { var _ = r[1]; }) };
        yield return new object[] { "this[\"A\"]", (Action<IRow>)(r => { var _ = r["A"]; }) };
        yield return new object[] { "Set(1, string)", (Action<IRow>)(r => r.Set(1, "x")) };
        yield return new object[] { "Set(1, double)", (Action<IRow>)(r => r.Set(1, 1.0)) };
        yield return new object[] { "Set(1, decimal)", (Action<IRow>)(r => r.Set(1, 1.0m)) };
        yield return new object[] { "Set(1, int)", (Action<IRow>)(r => r.Set(1, 1)) };
        yield return new object[] { "Set(1, long)", (Action<IRow>)(r => r.Set(1, 1L)) };
        yield return new object[] { "Set(1, bool)", (Action<IRow>)(r => r.Set(1, true)) };
        yield return new object[] { "Set(1, DateTime)", (Action<IRow>)(r => r.Set(1, DateTime.UnixEpoch)) };
        yield return new object[] { "Set(1, DateOnly)", (Action<IRow>)(r => r.Set(1, new DateOnly(2026, 1, 1))) };
        yield return new object[] { "Set(1, TimeOnly)", (Action<IRow>)(r => r.Set(1, new TimeOnly(12, 0))) };
        yield return new object[] { "Set(1, TimeSpan)", (Action<IRow>)(r => r.Set(1, TimeSpan.FromHours(1))) };
        yield return new object[] { "Underlying", (Action<IRow>)(r => { var _ = r.Underlying; }) };
        yield return new object[] { "Hidden get", (Action<IRow>)(r => { var _ = r.Hidden; }) };
        yield return new object[] { "Hidden set", (Action<IRow>)(r => { r.Hidden = true; }) };
    }

    [Theory]
    [MemberData(nameof(RowOperations))]
    public void Row_Member_Throws_ObjectDisposed_After_Workbook_Dispose(string memberName, Action<IRow> op)
    {
        var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var row = sheet.AppendRow();
        wb.Dispose();

        Action act = () => op(row);
        act.Should().Throw<ObjectDisposedException>(
            $"IRow.{memberName} must throw ObjectDisposedException once its owning workbook is disposed");
    }

    // ---- ICell ----------------------------------------------------------

    public static IEnumerable<object[]> CellOperations()
    {
        // Properties
        yield return new object[] { "Address", (Action<ICell>)(c => { var _ = c.Address; }) };
        yield return new object[] { "RowIndex", (Action<ICell>)(c => { var _ = c.RowIndex; }) };
        yield return new object[] { "ColumnIndex", (Action<ICell>)(c => { var _ = c.ColumnIndex; }) };
        yield return new object[] { "Kind", (Action<ICell>)(c => { var _ = c.Kind; }) };
        yield return new object[] { "Underlying", (Action<ICell>)(c => { var _ = c.Underlying; }) };
        // Setters
        yield return new object[] { "SetString", (Action<ICell>)(c => c.SetString("x")) };
        yield return new object[] { "SetRichText", (Action<ICell>)(c => c.SetRichText(new RichText(new RichTextRun("x")))) };
        yield return new object[] { "GetRichText", (Action<ICell>)(c => c.GetRichText()) };
        yield return new object[] { "SetNumber(double)", (Action<ICell>)(c => c.SetNumber(1.0)) };
        yield return new object[] { "SetNumber(decimal)", (Action<ICell>)(c => c.SetNumber(1.0m)) };
        yield return new object[] { "SetNumber(int)", (Action<ICell>)(c => c.SetNumber(1)) };
        yield return new object[] { "SetNumber(long)", (Action<ICell>)(c => c.SetNumber(1L)) };
        yield return new object[] { "SetBool", (Action<ICell>)(c => c.SetBool(true)) };
        yield return new object[] { "SetDate(DateTime)", (Action<ICell>)(c => c.SetDate(DateTime.UnixEpoch)) };
        yield return new object[] { "SetDate(DateOnly)", (Action<ICell>)(c => c.SetDate(new DateOnly(2026, 1, 1))) };
        yield return new object[] { "SetTime", (Action<ICell>)(c => c.SetTime(new TimeOnly(12, 0))) };
        yield return new object[] { "SetDuration", (Action<ICell>)(c => c.SetDuration(TimeSpan.FromHours(1))) };
        yield return new object[] { "SetFormula", (Action<ICell>)(c => c.SetFormula("=1+1")) };
        yield return new object[] { "Clear", (Action<ICell>)(c => c.Clear()) };
        // Getters
        yield return new object[] { "GetString", (Action<ICell>)(c => c.GetString()) };
        yield return new object[] { "GetNumber", (Action<ICell>)(c => c.GetNumber()) };
        yield return new object[] { "GetBool", (Action<ICell>)(c => c.GetBool()) };
        yield return new object[] { "GetDate", (Action<ICell>)(c => c.GetDate()) };
        yield return new object[] { "GetDateOnly", (Action<ICell>)(c => c.GetDateOnly()) };
        yield return new object[] { "GetTime", (Action<ICell>)(c => c.GetTime()) };
        yield return new object[] { "GetDuration", (Action<ICell>)(c => c.GetDuration()) };
        yield return new object[] { "GetError", (Action<ICell>)(c => c.GetError()) };
        yield return new object[] { "GetFormula", (Action<ICell>)(c => c.GetFormula()) };
        yield return new object[] { "Comment", (Action<ICell>)(c => c.Comment("x")) };
        yield return new object[] { "GetComment", (Action<ICell>)(c => c.GetComment()) };
        yield return new object[] { "GetCommentAuthor", (Action<ICell>)(c => c.GetCommentAuthor()) };
        yield return new object[] { "Hyperlink", (Action<ICell>)(c => c.Hyperlink("https://example.com")) };
        yield return new object[] { "GetHyperlink", (Action<ICell>)(c => c.GetHyperlink()) };
        yield return new object[] { "Style", (Action<ICell>)(c => c.Style(CellStyle.Default)) };
        yield return new object[] { "NumberFormat", (Action<ICell>)(c => c.NumberFormat("0.00")) };
        yield return new object[] { "GetStyle", (Action<ICell>)(c => c.GetStyle()) };
        yield return new object[] { "ApplyNamedStyle", (Action<ICell>)(c => c.ApplyNamedStyle("H")) };
    }

    [Theory]
    [MemberData(nameof(CellOperations))]
    public void Cell_Member_Throws_ObjectDisposed_After_Workbook_Dispose(string memberName, Action<ICell> op)
    {
        var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var cell = sheet["A1"];
        wb.Dispose();

        Action act = () => op(cell);
        act.Should().Throw<ObjectDisposedException>(
            $"ICell.{memberName} must throw ObjectDisposedException once its owning workbook is disposed");
    }

    // ---- IRange ---------------------------------------------------------

    public static IEnumerable<object[]> RangeOperations()
    {
        yield return new object[] { "Address", (Action<IRange>)(r => { var _ = r.Address; }) };
        yield return new object[] { "FirstRow", (Action<IRange>)(r => { var _ = r.FirstRow; }) };
        yield return new object[] { "LastRow", (Action<IRange>)(r => { var _ = r.LastRow; }) };
        yield return new object[] { "FirstCol", (Action<IRange>)(r => { var _ = r.FirstCol; }) };
        yield return new object[] { "LastCol", (Action<IRange>)(r => { var _ = r.LastCol; }) };
        yield return new object[] { "Count", (Action<IRange>)(r => { var _ = r.Count; }) };
        yield return new object[] { "Sheet", (Action<IRange>)(r => { var _ = r.Sheet; }) };
        yield return new object[] { "GetEnumerator", (Action<IRange>)(r =>
        {
            using var it = r.GetEnumerator();
            _ = it.MoveNext();
        }) };
        yield return new object[] { "EnumerateAll", (Action<IRange>)(r =>
        {
            using var it = r.EnumerateAll().GetEnumerator();
            _ = it.MoveNext();
        }) };
        yield return new object[] { "Value", (Action<IRange>)(r => r.Value("x")) };
        yield return new object[] { "Apply", (Action<IRange>)(r => r.Apply(CellStyle.Default)) };
        yield return new object[] { "ApplyNamedStyle", (Action<IRange>)(r => r.ApplyNamedStyle("H")) };
        yield return new object[] { "ClearContents", (Action<IRange>)(r => r.ClearContents()) };
        // Note: Merge() not in this matrix — Merge throws InvalidOperation
        // on overlap and a 1x1 is a no-op; the throw-after-dispose path
        // happens to coincide with ranges where Merge would already throw,
        // making the test fragile. The Range-level Merge dispatching to
        // sheet-level MergeCells is covered by RangeApiTests.
    }

    [Theory]
    [MemberData(nameof(RangeOperations))]
    public void Range_Member_Throws_ObjectDisposed_After_Workbook_Dispose(string memberName, Action<IRange> op)
    {
        var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var range = sheet.Range("A1:B2");
        wb.Dispose();

        Action act = () => op(range);
        act.Should().Throw<ObjectDisposedException>(
            $"IRange.{memberName} must throw ObjectDisposedException once its owning workbook is disposed");
    }

    // ---- IColumn --------------------------------------------------------

    public static IEnumerable<object[]> ColumnOperations()
    {
        yield return new object[] { "Index", (Action<IColumn>)(c => { var _ = c.Index; }) };
        yield return new object[] { "Letter", (Action<IColumn>)(c => { var _ = c.Letter; }) };
        yield return new object[] { "Sheet", (Action<IColumn>)(c => { var _ = c.Sheet; }) };
        yield return new object[] { "Hidden get", (Action<IColumn>)(c => { var _ = c.Hidden; }) };
        yield return new object[] { "Hidden set", (Action<IColumn>)(c => { c.Hidden = true; }) };
        yield return new object[] { "WidthUnits get", (Action<IColumn>)(c => { var _ = c.WidthUnits; }) };
        yield return new object[] { "WidthUnits set", (Action<IColumn>)(c => { c.WidthUnits = 10.0; }) };
        yield return new object[] { "Width", (Action<IColumn>)(c => c.Width(10.0)) };
        yield return new object[] { "AutoSize", (Action<IColumn>)(c => c.AutoSize()) };
        yield return new object[] { "ForEachPopulated", (Action<IColumn>)(c => c.ForEachPopulated(_ => { })) };
        yield return new object[] { "SetDefaultStyle", (Action<IColumn>)(c => c.SetDefaultStyle(CellStyle.Default)) };
    }

    [Theory]
    [MemberData(nameof(ColumnOperations))]
    public void Column_Member_Throws_ObjectDisposed_After_Workbook_Dispose(string memberName, Action<IColumn> op)
    {
        var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var col = sheet.Column("A");
        wb.Dispose();

        Action act = () => op(col);
        act.Should().Throw<ObjectDisposedException>(
            $"IColumn.{memberName} must throw ObjectDisposedException once its owning workbook is disposed");
    }

    // ---- Sanity: double-dispose is still a no-op (decision #42) --------

    [Fact]
    public void Double_Dispose_Across_Many_Configurations_Is_Safe()
    {
        for (int sheets = 0; sheets <= 3; sheets++)
        {
            var wb = Workbook.Create();
            for (int i = 0; i < sheets; i++) wb.AddSheet($"S{i}");
            wb.Dispose();
            Action second = () => wb.Dispose();
            second.Should().NotThrow($"second Dispose with {sheets} sheets");
        }
    }

    // Stub used to exercise the RemoveTable disposed-throw check —
    // ThrowIfDisposed fires before the foreign-table-rejection logic,
    // so the stub never gets dereferenced as anything but a non-null
    // reference.
    private sealed class XssfTableStub : ITable
    {
        public string Name => "";
        public string DisplayName { get => ""; set { } }
        public string Address => "";
        public ISheet Sheet => null!;
        public IReadOnlyList<string> ColumnNames => Array.Empty<string>();
        public bool HasTotalsRow => false;
        public string? StyleName { get => null; set { } }
        public NPOI.XSSF.UserModel.XSSFTable Underlying => null!;
        public void AddTotalsRow() { }
        public void RemoveTotalsRow() { }
        public void SetColumnTotal(string columnName, TotalsRowFunction function) { }
        public void SetColumnTotal(string columnName, string customFormula) { }
        public void SetColumnTotalLabel(string columnName, string label) { }
    }
}
