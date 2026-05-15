# Changelog

All notable changes to NetXlsx are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html)
per `docs/design.md §3 #23`. Pre-1.0 minor versions may include breaking
changes (decision I19).

## [Unreleased]

### Added
- Initial project scaffold (Directory.Build.props, Directory.Packages.props,
  nuget.config, .editorconfig, LICENSE, CODEOWNERS, build scripts,
  TeamCity DSL placeholder, source project skeletons, test/benchmark/
  sample/golden-file project skeletons, public-API snapshot files).
- Strong-name key (`netxlsx.snk`) generated and committed.
- All four pre-implementation spikes complete; results captured under
  `spikes/results/`. Three triggered design revisions:
  - Spike 4 (AOT/trim): both `PublishAot` and `PublishTrimmed` fail at
    runtime against NPOI 2.7.3. Roadmap matrix marked `No`; decision I2
    updated from "TBD pending" to measured outcome.
  - Spike 2 (streaming back-pressure): in-memory 100k-row target missed
    by ~2×. Target lowered to 30k rows; streaming recommended above that.
    Streaming sustained flat ~70 MB ΔGC at 500k rows.
  - Spike 1 (style dedup): "10%/30% overhead vs raw NPOI" framing was
    measuring a phantom — raw NPOI hits its style cap before completing
    any styled workload of meaningful size. Replaced with absolute
    capacity + throughput targets.
  - Spike 3 (async wrapping): `Task.Run` wrapping is free or net-positive
    at every size. Decision #5 stands without revision.

### Known scaffold placeholders (TODOs)
- `nuget.config` — internal feed URL pending.
- `Directory.Build.props` `RepositoryUrl` — github.com/jkindrix pending.
- `CODEOWNERS` — owning team identifiers pending.
- Source Link package is not wired; depends on github.com/jkindrix choice.

## [0.1.0] — TBD

The first tagged scaffold release. Cut once the placeholders above are
filled, the first green CI run completes, and the four pre-implementation
spikes (style-dedup, streaming back-pressure, async wrapping cost, AOT/trim
posture — see `docs/roadmap.md`) are scheduled or running.
