# NPOI Workarounds

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
