# NPOI Workarounds

> **Status update (2026-06-04): RETIRED — NPOI was removed from the library
> at the v2.0.0 cutover (decision I-82 in `docs/design.md`), so there is no
> engine left to work around.** The catalog closes with zero entries: the
> v1.x-era NPOI quirk handling lived inside the `Xssf*`/`Sxssf*` internals
> deleted at the cutover, and the S26 populate-as-discovered discipline was
> overtaken by the engine swap before any entry was recorded. The file is
> retained as the S26 record; do not add entries.

A catalog of NPOI bugs or behavioral quirks NetXlsx routes around, with
the workaround mechanism and the conditions under which the workaround
could be retired.

Per design decision **S26**, this file is populated as workarounds are
discovered. The discipline: every workaround written into the codebase has
a corresponding entry here, with enough detail that a future maintainer
can audit whether the workaround is still needed.

## Format per entry

```
## <Short title>

**NPOI version observed:** <version>
**Date observed:** <YYYY-MM-DD>
**Affected NPOI API:** <type/member>
**Symptom:** <what goes wrong>
**Workaround:** <what NetXlsx does instead, where in the codebase>
**Upstream link:** <issue URL, if reported>
**Retirement condition:** <which NPOI version or behavior change retires this>
```

## Entries

*(none yet — populated during implementation)*
