# Security Policy

## Reporting a vulnerability

Please **do not open a public GitHub issue** for security vulnerabilities.

Report them via [**private security advisories**](https://github.com/jkindrix/NetXlsx/security/advisories) on this repository. That route lets us discuss the issue and prepare a fix without exposing it before a patch is available.

If you don't have a GitHub account or can't use private advisories, email the maintainer (`jkindrix@gmail.com`) with `[NetXlsx security]` in the subject line.

## What to include

- A description of the issue and its impact.
- Steps to reproduce, ideally with a minimal example.
- The affected version(s) — git tag or commit hash if known.
- Whether you've shared this with anyone else.

## What to expect

- Acknowledgement within 5 business days.
- An initial assessment within 14 days.
- Coordinated disclosure on a timeline matched to severity — default is 90 days, sooner for critical issues.
- Credit in the release notes for the fix, unless you ask otherwise.

## Scope

Since v2.0.0 (decision I-82) the library's engine is Microsoft's
**Open XML SDK** (`DocumentFormat.OpenXml`) — its only runtime
dependency. Vulnerabilities specific to the SDK's OOXML parsing belong
upstream (https://github.com/dotnet/Open-XML-SDK); we'll coordinate with
its maintainers if a finding crosses the boundary.

### Dependency posture (read this if you parse untrusted input)

- **Runtime:** `DocumentFormat.OpenXml` is pinned exact and bumped
  deliberately via PR. It is MIT-licensed and actively maintained by
  Microsoft, so upstream security fixes flow through normal version
  bumps — the pre-v2.0.0 frozen-engine ceiling (NPOI pinned at 2.7.3
  for license reasons) no longer applies.
- **Malformed-input hardening:** the open path gates input explicitly
  (non-empty, workbook part present, exception classification into
  `MalformedFileException`), corrupt leaf values fail loud rather than
  silently defaulting (decision I-83), and the fuzz harness runs in CI.
- **Test-only NPOI:** NPOI 2.7.3 remains a dependency of the *test and
  benchmark projects only*, where it serves as an independent
  third-party reader/writer oracle over the engine's output. It never
  ships in the NetXlsx package. The `SixLabors.ImageSharp` /
  `System.Security.Cryptography.Xml` transitive pins in
  `Directory.Packages.props` exist solely to patch NPOI's vulnerable
  transitive closure for those non-shipping projects.

Out of scope:
- The strong-name signing key (`netxlsx.snk`) is committed in this repo by design. Strong-naming in OSS is a friction-reducer for legacy consumers, not a security boundary. "Anyone can sign as NetXlsx" is not a vulnerability — it's documented behavior.
- Issues in the Open XML SDK itself: please report to https://github.com/dotnet/Open-XML-SDK.
