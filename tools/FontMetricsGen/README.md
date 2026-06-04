# FontMetricsGen

One-off generator for `src/NetXlsx/Internal/FontMetricsTables.g.cs` — the
embedded font-metric tables behind `IColumn.AutoSize` on the Open XML SDK
engine (decision I-84 in `docs/design.md`).

## What it does

Parses the `head` / `hhea` / `maxp` / `cmap` / `hmtx` / `loca` / `glyf`
tables of a fixed set of TrueType fonts and emits per-character advance
widths (font units, `unitsPerEm`-relative), the ink width of the `'0'`
glyph (NPOI's `defaultCharWidth` divisor), and ascent/descent, as C#
source. **Only numbers are extracted — never font-file bytes or
outlines** — so NetXlsx's MIT licensing posture is unaffected.

## Provenance rules

- Sources must be **openly licensed** (SIL OFL). Today: Carlito
  (metric-compatible with Calibri) and Liberation Sans / Serif / Mono 2.x
  (metric-compatible with Arial / Times New Roman / Courier New).
- **Never extract from proprietary fonts** (e.g. the `msttcorefonts`
  Arial, Windows-mounted `Calibri`), even though only numbers ship —
  the open metric-twin exists precisely so provenance stays clean.
- Liberation **1.x** is excluded deliberately (GPL+exception lineage);
  that is why there is no Arial Narrow table — `AutoSize` on an
  Arial-Narrow cell throws `MissingFontException` by design until an
  OFL metric twin is sourced.

## Regenerating

```bash
# Debian/Ubuntu: apt-get install fonts-crosextra-carlito fonts-liberation2
dotnet run --project tools/FontMetricsGen
# optional: dotnet run --project tools/FontMetricsGen -- <fontsRoot> <outputPath>
```

The emitted file header records each source file's sha256 and full font
name. Re-run only when adding a metric family or when the distro font
packages meaningfully change; commit the regenerated file together with
the matching resolver change in `src/NetXlsx/Internal/OoxmlFontMetrics.cs`.

This project is intentionally **not** in `NetXlsx.sln` — it is a
maintenance tool, not a build step; the generated tables are committed.
