# Changelog

Sourced from actual commit history (`git log`). This covers the substantive changes to the build script, workflow, and site; the ~100 routine "Publish Daily Prophet issue" auto-commits from the twice-daily scheduled job are omitted here for readability but represent continuous daily operation since 2026-03-28.

## 2026-03-28 — Initial build
- Initialized the Daily Prophet Pages site and rendered the first morning issue.
- Aligned the formatter to the target template spec and fixed initial rendering bugs.
- Added the scheduled publishing GitHub Actions workflow, then hardened it and fixed a change-detection bug.
- Archived the first issue (`20260328-AM`).

## 2026-03-29 — First late/breaking edition
- Published a combined late-morning edition and a breaking-news edition (`20260329-AM-late`, `20260329-BN-garden-os-sprint0-green`).

## 2026-03-30 — Archive and resilience fixes
- Fixed the workflow to commit the `issues/` archive alongside `index.html` (previously only the current issue was persisted).
- Hardened the build script against malformed or incomplete Notion pages.

## 2026-04-06 — Notion pipeline hardening
- Fixed a null-array crash caused by PowerShell's automatic pipeline unwrapping.
- Fixed pipeline enumeration broken by `return ,$array` semantics.
- Added validation and error handling for Notion block IDs.

## 2026-06-20 — 2026-06-21 — Favicon
- Added a site favicon (in response to a favicon audit).
- Fixed the favicon not persisting across the daily auto-publish commits.

## 2026-07-06 — Licensing
- Added an explicit all-rights-reserved `LICENSE`.

## 2026-08-26 — Audit remediation (2026-08-24 portfolio audit H-1, H-2, M-1, M-2)
- H-1: the front page now shows the day's most recently created edition (Morning, Evening, or Breaking) instead of only Morning, and the workflow gained 6:05 PM ET evening runs so the twice-daily cadence actually publishes the PM briefing.
- H-2: archive files are materialized for every database page on both the normal and fallback paths, so the archive ledger can no longer link to files that were never written (15 ledger links were 404 on the live site).
- M-1: added `validate-daily-prophet.ps1` and a workflow step that fails the publish on an incomplete front page, unresolved archive ledger links, unreplaced template slots, or a fallback front page while today's issue is archived.
- M-2: the generated issue and archive pages now carry header/main/footer/nav landmarks and an h1/h2/h3 heading outline for screen-reader navigation, with unchanged visual output.

## Ongoing
- Twice-daily automated "Publish Daily Prophet issue" commits from the scheduled workflow, running continuously from 2026-03-28 through the present, each regenerating `index.html` and archiving the issue under `issues/`.
