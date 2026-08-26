[CmdletBinding()]
param()

# Post-build content validation for the Daily Prophet site (audit M-1).
# Fails the publish when the rendered output is structurally broken:
#   1. index.html missing, incomplete, or carrying unreplaced template slots
#   2. archive ledger links pointing at files that do not exist
#   3. the "No issue yet" fallback front page while today's issue is archived
#   4. archived issues carrying unreplaced template slots

$ErrorActionPreference = 'Stop'

$RepoRoot = $PSScriptRoot
$IndexPath = Join-Path $RepoRoot 'index.html'
$IssuesDirectory = Join-Path $RepoRoot 'issues'
$ArchiveIndexPath = Join-Path $IssuesDirectory 'index.html'

function Get-EnvValue([string]$Name) {
  $processValue = [Environment]::GetEnvironmentVariable($Name, 'Process')
  if (-not [string]::IsNullOrWhiteSpace($processValue)) { return $processValue }

  $userValue = [Environment]::GetEnvironmentVariable($Name, 'User')
  if (-not [string]::IsNullOrWhiteSpace($userValue)) { return $userValue }

  return $null
}

function Get-IssueDate() {
  $override = Get-EnvValue 'ISSUE_DATE'
  if ($override) { return $override }

  $utcNow = (Get-Date).ToUniversalTime()
  $eastern = [System.TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time')
  return [System.TimeZoneInfo]::ConvertTimeFromUtc($utcNow, $eastern).ToString('yyyy-MM-dd')
}

$failures = @()
$indexHtml = $null

# 1. Front page must exist and be a complete render.
if (-not (Test-Path -LiteralPath $IndexPath)) {
  $failures += 'index.html is missing.'
} else {
  $indexHtml = Get-Content -LiteralPath $IndexPath -Raw
  if ($indexHtml -notmatch '<h1 class="mh-title">') {
    $failures += 'index.html is missing the masthead heading.'
  }
  if ($indexHtml -match '\{\{[A-Z_]+\}\}') {
    $failures += 'index.html contains unreplaced template slots.'
  }
}

# 2. Every archive ledger link must resolve to a real file.
$ledgerLinks = @()
if (-not (Test-Path -LiteralPath $ArchiveIndexPath)) {
  $failures += 'issues/index.html (archive ledger) is missing.'
} else {
  $archiveHtml = Get-Content -LiteralPath $ArchiveIndexPath -Raw
  $ledgerLinks = @([regex]::Matches($archiveHtml, 'class="archive-link" href="([^"]+)"') |
    ForEach-Object { [System.Net.WebUtility]::HtmlDecode($_.Groups[1].Value) })
  foreach ($link in $ledgerLinks) {
    if (-not (Test-Path -LiteralPath (Join-Path $IssuesDirectory $link))) {
      $failures += ('Archive ledger links to a missing file: issues/' + $link)
    }
  }
}

# 3. The fallback front page is only acceptable when today truly has no issue.
$todayCompact = (Get-IssueDate) -replace '-', ''
$todayLinks = @($ledgerLinks | Where-Object { $_ -like ($todayCompact + '*') })
if ($todayLinks.Count -gt 0 -and $indexHtml -and $indexHtml -match 'class="lead-hl">No issue yet</') {
  $failures += ("index.html is the 'No issue yet' fallback although today's issue is archived: " + ($todayLinks -join ', '))
}

# 4. Archived issues must be complete renders.
if (Test-Path -LiteralPath $IssuesDirectory) {
  foreach ($file in Get-ChildItem -LiteralPath $IssuesDirectory -Filter '*.html') {
    $issueHtml = Get-Content -LiteralPath $file.FullName -Raw
    if ($issueHtml -match '\{\{[A-Z_]+\}\}') {
      $failures += ('issues/' + $file.Name + ' contains unreplaced template slots.')
    }
  }
}

if ($failures.Count -gt 0) {
  foreach ($failure in $failures) {
    Write-Host ('FAIL: ' + $failure)
  }
  throw ('Daily Prophet validation failed with ' + $failures.Count + ' problem(s).')
}

Write-Host ('Validation passed: front page OK, ' + $ledgerLinks.Count + ' ledger link(s) resolve.')
