# The Daily Prophet

A wizarding-newspaper-themed personal briefing page, published on demand from a Notion database and served as a static GitHub Pages site.

When dispatched manually, GitHub Actions pulls the selected day's briefing from Notion and regenerates `index.html`. The issue includes a front page, a "Hogwarts Watchlist" of active projects, potion notes, spells, a map note, and an optional translated briefing. Each rendered issue is also archived under `issues/` so past editions stay browsable.

## What's here

| Path | What it is |
|---|---|
| `index.html` | The current issue — the page GitHub Pages serves at the repo root |
| `build-daily-prophet.ps1` | PowerShell script that queries the Notion API, parses the day's page into sections, and regenerates `index.html` plus an archive copy |
| `issues/` | Archived past issues (`YYYYMMDD-AM.html`, `YYYYMMDD-PM.html`, etc.) and an `index.html` listing them |
| `validate-daily-prophet.ps1` | Post-build validation — fails the publish if the front page is incomplete, an archive ledger link points at a missing file, or the fallback page ships while today's issue exists |
| `.github/workflows/publish-daily-prophet.yml` | Manual workflow that runs the build script, validates the rendered output, and commits and pushes changed output |
| `favicon.svg` | Site favicon |

## How to run it

There's no app to install — this is a PowerShell build script plus a static HTML output.

- **Publish from GitHub Actions:** choose **Publish Daily Prophet** and select **Run workflow** (`workflow_dispatch`). Optionally enter `issue_date` in `YYYY-MM-DD` format to build a specific day. Scheduled publication is paused.
- **Run locally:**
  ```powershell
  $env:NOTION_API_KEY = "<your Notion integration token>"
  ./build-daily-prophet.ps1
  ```
  This regenerates `index.html` and writes/updates the matching file(s) in `issues/`. An optional `ISSUE_DATE` environment variable overrides the date (defaults to "now" in US/Eastern).
- **View the output:** open `index.html` directly, or visit the page as served by GitHub Pages.

## Conventions

- The script targets a specific Notion database (`$DatabaseId` in `build-daily-prophet.ps1`) with a `Date`/`Issue ID`-sorted schema; page content is parsed by convention (bold rich-text segments become headlines, page mentions become sourced attributions, callout blocks become watchlist cards).
- Output files are written UTF-8 without BOM with normalized CRLF line endings (`Write-Utf8File`), matching the `.gitattributes` rule forcing CRLF for `*.ps1`/`*.html`.
- The workflow only commits and pushes when the rendered output actually changed, so a Notion query returning identical content produces no-op runs.
