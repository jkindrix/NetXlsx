# Golden-File Fixtures

Per design `I18` and roadmap *Process rules → Test fixture provenance*, every
fixture in this directory must have a sibling `<fixture-name>.fixture.md`
describing its provenance.

## Required metadata per fixture

- **Source** — hand-crafted in Excel (which version, which OS), or produced
  by `<fixture-name>.gen.cs`.
- **Purpose** — which golden-file test or recipe uses this fixture.
- **Sensitive content** — explicit "none" or list of sanitization performed.
  No fixture may contain real customer data, real account numbers, or real
  PII. Synthetic sample data is acceptable.
- **Author**, **date added**.

## Fixtures expected before v1.0 ship

(Each row from `docs/design.md §8.1` cookbook list, plus the round-trip
preservation fixture that gates v1.0.)

- `hello-workbook.xlsx`
- `tabular-export.xlsx`
- `typed-export.xlsx`
- `typed-import.xlsx`
- `styled-report.xlsx`
- `formulas.xlsx`
- `multi-sheet.xlsx`
- `hyperlinks-and-comments.xlsx`
- `streaming-million-rows.xlsx`
- `npoi-escape-hatch.xlsx`
- `open-edit-save-roundtrip.xlsx`  *(v1.0 ship-blocker — see §7.7)*
- `time-and-duration.xlsx`
- `cell-errors.xlsx`

This directory is otherwise empty in scaffold.
