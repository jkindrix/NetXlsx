<!-- Thanks! CONTRIBUTING.md is short and load-bearing — the checklist
     below is its enforcement surface. -->

## What & why

<!-- What changed and why. Reference a design decision (docs/design.md
     I-NN / #N) for any public-surface change, or motivate a new one. -->

## Checklist

- [ ] Conventional commit subjects (`feat:` / `fix:` / `docs:` / …), no AI-tool branding
- [ ] Behavior change ⇒ tests change (`bash build/build.sh test` green on net8.0 + net10.0)
- [ ] New/changed public API ⇒ `PublicAPI.Unshipped.txt` entry (RS0016 enforces) and XML docs
- [ ] New/changed generator diagnostic ⇒ `AnalyzerReleases.Unshipped.md` entry
- [ ] Emitted-bytes change ⇒ golden/round-trip expectations regenerated deliberately and called out below
- [ ] `CHANGELOG.md` entry under `[Unreleased]`
