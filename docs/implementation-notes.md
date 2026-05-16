# NetXlsx — Implementation Notes

A working file capturing patterns and lessons from the implementation
phase, with the project context they emerged in. This is **not** a
methodology — methodologies are written after the practices they describe
have been validated by completion. This file is the raw material from
which an implementation-phase methodology can be distilled after v1.0
ships.

Entries are dated and project-specific. Generalize at your own risk.

---

## 2026-05-15 — From v0.1.0 design lock to v0.2.0 first round-trip

### Vertical slice scoping

The reviewer pushed for a thin v0.2.0 vertical slice (Create → AddSheet →
SetString on `A1` → Save → Open → GetString) rather than building the full
§6 surface in one step. We did it. Outcome:

- The slice surfaced two genuine design questions the spec had not pinned
  down: how `XSSFWorkbook.Write` handles stream lifetime (NPOI closes by
  default; we override) and the exact form of `[MaybeNullWhen]` in the
  public-API analyzer file. Both were absorbed in 5-line fixes.
- 63 tests passing, full round-trip working, in one focused commit.
- Compared to the alternative ("ship the full §6"), the slice gave the
  source generator's `[Obsolete]` decoration a real reason to exist — the
  generator's stub methods still throw because `ISheet.AppendRow` etc.
  haven't landed yet, which they wouldn't have in either approach.

**Pattern:** when the design surface is large, scope the first
implementation milestone to an end-to-end thread, not a horizontal slice.
The thread surfaces specification gaps that horizontal work cannot.

### Stub policy: throw vs `[Obsolete(error: true)]` vs absent

Three options for an API surface that exists in spec but isn't yet wired:

| Option | Failure surface | When to use |
|---|---|---|
| Absent (don't emit) | Compile error: method doesn't exist | When the absence is reversible and no consumer should depend on it yet |
| `throw new NotImplementedException` | Runtime crash | Almost never — silently compiles, fails late, very poor signal |
| `[Obsolete(error: true)]` on the method | Compile error CS0619 with explanatory message | When the method *signature* is committed and consumers should write code against it that breaks at the right time |

We landed on `[Obsolete(error: true)]` for the generator's emitted
extensions. Reviewer specifically called out the throwing-stub anti-pattern
in commit `6af2ea8`. Lesson: pick the loudest failure your consumer-facing
contract can support.

### Design-doc revisions during implementation

When implementation surfaced something the design hadn't anticipated, we
took one of three paths:

1. **Edit the design row and continue.** Used when the surfaced item was
   small enough that the rationale fit in a sentence (e.g., the `Underlying`
   property type for `IWorkbook` resolving to `XSSFWorkbook` concretely).
2. **Treat as a structural revision and pause.** Used when a spike result
   invalidated an architectural assumption (spike 4: AOT/trim
   incompatibility re-classified two roadmap rows from `TBD*` to `No`).
3. **Defer with explicit milestone marker.** Used when the gap was real
   but the slice's scope didn't include it (concurrent-mutation detection
   per decision #43 — documented as known limitation in v0.2.0 CHANGELOG).

The discipline: never the fourth option, which is silent default. Every
implementation-time decision either edits the design or is recorded in
CHANGELOG as a deferred item.

### The PublicAPI analyzer earns its keep

`Microsoft.CodeAnalysis.PublicApiAnalyzers` blocked compile every time a
public symbol was added without a matching entry in
`PublicAPI.Unshipped.txt`. This forced the author to *think* about the
public surface in the same commit that introduced it — couldn't drift.
The wider lesson: a quality gate that the toolchain enforces beats a
quality gate that the reviewer enforces.

### Review-during-implementation cadence

After each milestone commit, we ran the work through an external review.
Patterns observed:

- The first review of any new code pass is dense — substantive items.
- The second review on the same code is usually cosmetic.
- Reviewers occasionally flag false alarms based on file listings (bin/,
  TestResults/) that are actually gitignored. Verify with `git ls-files`
  before acting. (Now codified in design-phase-methodology.md §5.)

### Honest constraint admission > optimism

The AOT/trim spike (spike 4) was the most valuable single thing we did
this session. The original design row was "TBD pending spike, likely
trim-with-warnings, AOT-incompatible." Spike 4 measured: both AOT and
trim fail at runtime, full stop. The roadmap matrix now says `No†` with
a real footnote. Consumers who would have hit a runtime crash six months
into adoption now get a build-time error courtesy of
`buildTransitive/NetXlsx.targets`.

### Analyzer noise vs analyzer signal

`TreatWarningsAsErrors=true` plus `latest-recommended` analysis level
produced a dozen analyzer errors during the v0.2.0 work — most of which
were genuinely useful (`ArgumentNullException.ThrowIfNull`,
`ObjectDisposedException.ThrowIf`, `string.Contains(char)` over
`IndexOf`). Two were noise for our specific design (`CA1720` on enum
value `String`; `RS0026` on intentional optional-param overloads). Both
were suppressed in `.editorconfig` with explicit comments.

**Pattern:** keep the gate; suppress at the rule level (not the file
level) when the rule is wrong for your design and the suppression has a
written justification.

---

## Future entries

Add a dated section per substantive implementation milestone. After
v1.0 ships, distill the patterns that recurred into a sibling methodology
document under `~/dev/projects/references/project-methodologies/`.
