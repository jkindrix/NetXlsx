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

The library wraps NPOI 2.7.x. Vulnerabilities specific to NPOI's OOXML parsing belong upstream (https://github.com/nissl-lab/npoi); we'll coordinate with NPOI maintainers if a finding crosses the boundary.

Out of scope:
- The strong-name signing key (`netxlsx.snk`) is committed in this repo by design. Strong-naming in OSS is a friction-reducer for legacy consumers, not a security boundary. "Anyone can sign as NetXlsx" is not a vulnerability — it's documented behavior.
- Issues in NPOI itself: please report to https://github.com/nissl-lab/npoi.
