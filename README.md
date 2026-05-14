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

sheet[2, 1].Value("North");                          // row 2, col 1 == A2 (1-indexed)
sheet[2, 2].Value(1234.56m).NumberFormat("$#,##0.00");

sheet.Column("B").Width(18);
sheet.FreezeRows(1);

await wb.SaveAsync("sales.xlsx");
```

Integer cell coordinates are **1-indexed** to match Excel's UI (`sheet[1, 1]` is `A1`).
