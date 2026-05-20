# NPOI 3.x adoption plan

NetXlsx is pinned at NPOI 2.7.3 per decision **I23** (last clean
Apache-2.0 release before the Open Source Maintenance Fee EULA
introduced in 2.8.0). When and if NPOI 3.x ships, this doc is the
playbook for evaluating and (if accepted) adopting it.

The longer-horizon "leave NPOI entirely and implement OOXML
ourselves" path lives in `docs/long-term.md` + `docs/v2-ooxml-planning.md`.
These two docs are complementary — one says "what we do if upstream
stabilizes," the other says "what we do if we decide to leave."

## Triggers — when this plan activates

Any one of:

1. NPOI 3.x is released with an explicitly OSI-approved license
   (Apache-2.0 or similar, no OSMF EULA on the binary).
2. The OSMF terms are clarified to exclude transitive consumers
   (i.e., libraries that wrap NPOI don't transitively obligate
   their own consumers).
3. A community fork emerges that maintains the 2.7.x line under
   pure Apache-2.0 with active CVE servicing.

The quarterly **Spike 5-Q** in `docs/scheduled-spikes.md` is where
each of these gets checked.

## Scope to evaluate

Per the NetXlsx surface, the NPOI 3.x acceptance criteria are:

### Required (no NPOI 3.x bump until all satisfied)

- [ ] **Apache-2.0 or equivalent OSI license** with no EULA on
      binary releases.
- [ ] **API compatibility for our consumed surface.** The full list:
  - `XSSFWorkbook`, `XSSFSheet`, `XSSFRow`, `XSSFCell`, `XSSFColumn`
    constructors and the members we exercise.
  - `XSSFColor` construction (we already use `CT_Color`-based
    construction — forward-compat fix from commit `42fbda3`).
  - `SXSSFWorkbook`, `SXSSFSheet`, `SXSSFRow`, `SXSSFCell` for the
    streaming path.
  - `CellStyle` / `IndexedColorMap` / `XSSFCellStyle` for the style
    pool.
  - `IConditionalFormatting` for the preservation test fixture.
  - `OPCPackage` + `PackagePart` for the raw OPC escape hatch.
  - `XSSFHyperlink` / `IComment` / `IName` for the annotation +
    named-range surface.
- [ ] **AOT + trim compatibility** measured via a re-run of
      `spikes/NetXlsx.AotSpike/`. If either still fails at runtime,
      the MSBuild guards `NXLS0100/0101` stay in place and the
      matrix rows in `roadmap.md` stay `Deferred†`.
- [ ] **No regression on the current test suite** (434 per TFM at
      v1.0.0; future releases may add more — assert "current count,
      no failures" rather than a literal number).
- [ ] **No regression > 15% on any benchmark in
      `benchmarks/NetXlsx.Benchmarks/`** (the same threshold our CI
      gate enforces).

### Welcome but not blocking

- AOT cleanness even if not yet officially supported.
- Improvements to the workbook-name-uniqueness constraint
  (NPOI 2.7.x rejects same-text workbook + sheet-scope name
  coexistence even though Excel itself permits it — documented in
  `implementation-notes.md`).
- A way to write 1904-epoch workbooks without escape-hatch
  gymnastics (NPOI 2.7.x hardcodes `date1904 = false`).
- Removal of the `XSSFColor(byte[])` ctor in 2.7.4 / `CT_Color`
  ctor `[Obsolete]` in 2.7.6 being resolved in 3.x with a stable
  API.

## Migration playbook (if accepted)

When the required criteria above are all green:

### Step 1 — preflight (1-2 days)

- Create a branch `npoi-3x-spike`.
- Bump `Directory.Packages.props`: `NPOI` 2.7.3 → 3.x.
- Try a clean build. Catalog every compile error.
- For each compile error, identify the API change and the
  replacement. **Do not fix yet — catalog only.** Output: a list
  of touch sites with proposed replacements.

### Step 2 — decide forward or back (1 day)

- If the catalog is < ~20 sites with mechanical replacements:
  proceed to step 3.
- If the catalog is > 20 sites or any replacement requires
  semantic rework (e.g., changes to our public contract):
  branch the decision back to the project owner. The cost may
  exceed the OOXML-from-scratch path described in
  `docs/v2-ooxml-planning.md`.

### Step 3 — apply replacements (3-5 days)

- One replacement per commit when feasible, so blame stays clean.
- Run the full test matrix after each commit.
- Document any new NPOI quirks in `docs/npoi-workarounds.md`.

### Step 4 — re-run spikes (1 day)

- Spike 4-Q: AOT + trim posture. If either now succeeds, lift the
  matrix row in `roadmap.md` and update decision **I2**.
- Spike 5-Q: re-verify the license posture hasn't changed since
  the trigger event.
- Benchmark suite: ensure no benchmark regresses > 15% vs the
  pre-bump baseline. Refresh the cached baseline if needed.

### Step 5 — release (1 day)

- Squash the spike branch onto main via a single PR.
- Bump the NetXlsx minor version (e.g., v1.0 → v1.1 if NPOI 3.x
  is non-breaking for our consumers; v2.0 if it is).
- New decision entry in `docs/design.md`: "I-NN — NPOI 3.x
  adopted on YYYY-MM-DD; supersedes I23."
- CHANGELOG entry with full migration impact.
- Update transitive-dep overrides in `Directory.Packages.props`:
  some of the CVE pins we carry (SixLabors.ImageSharp,
  System.Security.Cryptography.Xml) may become redundant once
  NPOI 3.x picks up the patched transitives directly.

## What this plan does NOT do

- **Pre-commit to adopting NPOI 3.x.** The OSMF situation, the
  NPOI 2.7.x patch churn (breaking API changes in 2.7.4 and 2.7.6
  patches), and the project's overall direction may all combine
  to argue *against* adoption. The plan codifies *how* we'd adopt
  if we choose to.
- **Replace `docs/long-term.md`.** Adopting NPOI 3.x is the v1.x
  path; implementing OOXML directly is the v2.0 path. They are
  alternatives — taking the NPOI 3.x bump doesn't close the
  v2.0 R&D track, and starting the v2.0 R&D doesn't preclude
  taking NPOI 3.x in the interim.

## Status

- **Current** (2026-05-20): No NPOI 3.x release. None of the
  three triggers have fired. Plan dormant; quarterly Spike 5-Q
  is the next check.
