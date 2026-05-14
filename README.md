# NetXlsx

Idiomatic C# facade over [NPOI](https://github.com/nissl-lab/npoi) for creating and reading `.xlsx` spreadsheets.

**Status:** Pre-implementation. Design phase.

## Documentation

- [Design document](docs/design.md) — foundational decisions, interface sketch, performance targets, quality gates.
- [Roadmap](docs/roadmap.md) — binary feature-scope matrix across releases, per-release progress checklists.

## At a glance

```csharp
using var wb = Workbook.Create();
var sheet = wb.AddSheet("Sales");

sheet["A1"].Value("Region").Bold();
sheet["B1"].Value("Revenue").Bold();

sheet["A2"].Value("North");
sheet["B2"].Value(1234.56m).NumberFormat("$#,##0.00");

sheet.Column("B").Width(18);
sheet.FreezeRows(1);

await wb.SaveAsync("sales.xlsx");
```
