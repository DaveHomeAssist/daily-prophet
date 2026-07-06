# The Daily Prophet

A wizarding-newspaper-themed personal briefing page, auto-published twice a day from a Notion database and served as a static GitHub Pages site.

Twice daily, a scheduled GitHub Actions workflow pulls the latest entry from a Notion database and re-renders `index.html` — a Harry Potter *Daily Prophet* pastiche with a front page, a "Hogwarts Watchlist" of active projects, potion notes, spells, a map note, and an optional translated briefing. Each rendered issue is also archived under `issues/` so past editions stay browsable.

## What's here

| Path | What it is |
|---|---|
| `index.html` | The current issue — the page GitHub Pages serves at the repo root |
| `build-daily-prophet.ps1` | PowerShell script that queries the Notion API, parses the day's page into sections, and regenerates `index.html` plus an archive copy |
| `issues/` | Archived past issues (`YYYYMMDD-AM.html`, `YYYYMMDD-PM.html`, etc.) and an `index.html` listing them |
| `.github/workflows/publish-daily-prophet.yml` | Scheduled workflow — runs the build script twice daily (and on manual dispatch with an optional `issue_date` override), commits, and pushes if the output changed |
| `favicon.svg` | Site favicon |

## How to run it

There's no app to install — this is a PowerShell build script plus a static HTML output.

- **Automatic (normal path):** the GitHub Actions workflow runs on its cron schedule and does everything (build, commit, push). No local action needed.
- **Manual re-run:** trigger the workflow by hand from the Actions tab (`workflow_dispatch`), optionally passing an `issue_date` (`YYYY-MM-DD`) to rebuild a specific day.
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
