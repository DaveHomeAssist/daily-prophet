[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$DatabaseId = '331255fc8f4480d59f92ffb703398031'
$OutputPath = Join-Path $PSScriptRoot 'index.html'
$IssuesDirectory = Join-Path $PSScriptRoot 'issues'
$ArchiveIndexPath = Join-Path $IssuesDirectory 'index.html'
$EmDash = [string][char]0x2014
$EnDash = [string][char]0x2013
$DashPattern = '\s*(' + [regex]::Escape($EmDash) + '|' + [regex]::Escape($EnDash) + '|-)\s*'

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

function Get-DateDisplay([string]$IssueDate) {
  return [datetime]::ParseExact($IssueDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture).ToString('dddd, d MMMM yyyy')
}

function Get-NotionHeaders() {
  $token = Get-EnvValue 'NOTION_API_KEY'
  if (-not $token) {
    throw 'NOTION_API_KEY is missing.'
  }

  return @{
    'Authorization'  = "Bearer $token"
    'Notion-Version' = '2022-06-28'
    'Content-Type'   = 'application/json'
  }
}

function Invoke-NotionJson([string]$Method, [string]$Uri, $Body = $null) {
  $params = @{
    Method  = $Method
    Uri     = $Uri
    Headers = Get-NotionHeaders
  }

  if ($null -ne $Body) {
    $params.Body = ($Body | ConvertTo-Json -Depth 30)
  }

  return Invoke-RestMethod @params
}

function Invoke-DatabaseQuery($Body) {
  return Invoke-NotionJson -Method Post -Uri "https://api.notion.com/v1/databases/$DatabaseId/query" -Body $Body
}

function Get-NotionChildren([string]$BlockId) {
  $results = @()
  $cursor = $null

  do {
    $uri = "https://api.notion.com/v1/blocks/$($BlockId -replace '-','')/children?page_size=100"
    if ($cursor) {
      $uri += "&start_cursor=$cursor"
    }

    $response = Invoke-NotionJson -Method Get -Uri $uri
    if ($response.results) {
      $results += @($response.results)
    }
    $cursor = $response.next_cursor
  } while ($cursor)

  return @($results)
}

function Get-AllIssuePages() {
  $results = @()
  $cursor = $null

  do {
    $body = @{
      sorts = @(
        @{ property = 'Date'; direction = 'descending' },
        @{ property = 'Issue ID'; direction = 'descending' }
      )
      page_size = 100
    }

    if ($cursor) {
      $body.start_cursor = $cursor
    }

    $response = Invoke-DatabaseQuery $body
    if ($response.results) {
      $results += @($response.results)
    }
    $cursor = $response.next_cursor
  } while ($cursor)

  return @($results)
}

function HtmlEncode([string]$Text) {
  if ($null -eq $Text) { return '' }
  return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Write-Utf8File([string]$Path, [string]$Content) {
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  $normalized = [regex]::Replace($Content, "`r?`n", "`r`n")
  [System.IO.File]::WriteAllText($Path, $normalized, $utf8NoBom)
}

$MentionCache = @{}

function Get-MentionTitle([string]$PageId) {
  if ([string]::IsNullOrWhiteSpace($PageId)) { return '' }
  if ($MentionCache.ContainsKey($PageId)) { return $MentionCache[$PageId] }

  try {
    $page = Invoke-RestMethod -Method Get -Uri "https://api.notion.com/v1/pages/$($PageId -replace '-','')" -Headers (Get-NotionHeaders) -TimeoutSec 10
    $titleProperty = $page.properties.PSObject.Properties | Where-Object { $_.Value.type -eq 'title' } | Select-Object -First 1
    $title = ''
    if ($titleProperty) {
      $title = (($titleProperty.Value.title | ForEach-Object { $_.plain_text }) -join '')
    }
    $MentionCache[$PageId] = $title
    return $title
  } catch {
    $MentionCache[$PageId] = ''
    return ''
  }
}

function Get-RichSegments($RichText) {
  $segments = @()

  foreach ($item in @($RichText)) {
    if ($item.type -eq 'mention' -and $item.mention.type -eq 'page') {
      $title = Get-MentionTitle $item.mention.page.id
      if ($title) {
        $segments += [pscustomobject]@{
          Text      = $title
          Bold      = $false
          IsMention = $true
        }
      }
      continue
    }

    $segments += [pscustomobject]@{
      Text      = $item.plain_text
      Bold      = [bool]$item.annotations.bold
      IsMention = $false
    }
  }

  return @($segments)
}

function Clean-Separator([string]$Text) {
  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

  $cleaned = $Text
  $cleaned = $cleaned -replace ('\s*(' + [regex]::Escape($EmDash) + '|' + [regex]::Escape($EnDash) + '|-)+\s*$'), ''
  $cleaned = $cleaned -replace ('^\s*(' + [regex]::Escape($EmDash) + '|' + [regex]::Escape($EnDash) + '|-)+\s*'), ''
  $cleaned = $cleaned -replace '\s{2,}', ' '
  return $cleaned.Trim()
}

function Join-Segments($Segments, [scriptblock]$Filter) {
  $parts = foreach ($segment in @($Segments)) {
    if (& $Filter $segment) {
      $segment.Text
    }
  }

  return Clean-Separator (($parts -join ' ').Trim())
}

function Get-PlainFromRichText($RichText) {
  $segments = Get-RichSegments $RichText
  return Clean-Separator (($segments | ForEach-Object { $_.Text }) -join ' ')
}

function Get-BlockPlainText($Block) {
  if (-not $Block) { return '' }

  $payload = $Block.$($Block.type)
  if (-not $payload) { return '' }

  if ($payload.PSObject.Properties.Name -contains 'rich_text') {
    return Get-PlainFromRichText $payload.rich_text
  }

  return ''
}

function Get-RecursiveBlockText($Block) {
  $parts = @()

  $ownText = Get-BlockPlainText $Block
  if ($ownText) {
    $parts += $ownText
  }

  if ($Block.has_children) {
    foreach ($child in (Get-NotionChildren $Block.id)) {
      $childText = Get-RecursiveBlockText $child
      if ($childText) {
        $parts += $childText
      }
    }
  }

  return Clean-Separator (($parts -join ' ').Trim())
}

function Split-RichTextRecord($RichText) {
  $segments = Get-RichSegments $RichText
  $headline = Join-Segments $segments { param($s) $s.Bold -and -not $s.IsMention }
  $summary  = Join-Segments $segments { param($s) (-not $s.Bold) -and (-not $s.IsMention) }
  $source   = Join-Segments $segments { param($s) $s.IsMention }

  if (-not $headline) {
    $fullText = Clean-Separator (($segments | ForEach-Object { $_.Text }) -join ' ')
    $parts = [regex]::Split($fullText, $DashPattern, 2)
    $headline = if ($parts.Count -ge 1) { $parts[0].Trim() } else { $fullText }
    $summary = if ($parts.Count -ge 2) { Clean-Separator $parts[1] } else { '' }
  }

  return [pscustomobject]@{
    Headline = $headline
    Summary  = $summary
    Source   = $source
  }
}

function Get-ColumnChildren([string]$ColumnListId) {
  $columns = @()
  foreach ($column in (Get-NotionChildren $ColumnListId)) {
    $columns += ,@(Get-NotionChildren $column.id)
  }
  return @($columns)
}

function Get-RibbonIcon([string]$EditionLabel) {
  switch ($EditionLabel) {
    'Morning Edition' { return '&#9728;' }
    'Evening Edition' { return '&#127769;' }
    'Breaking News'   { return '&#9889;' }
    default           { return '&#9728;' }
  }
}

function Get-EditionLabel([string]$EditionValue) {
  switch ($EditionValue) {
    'Morning'  { return 'Morning Edition' }
    'Evening'  { return 'Evening Edition' }
    'Breaking' { return 'Breaking News' }
    'Breaking News' { return 'Breaking News' }
    default    { return $EditionValue }
  }
}

function Get-PropertyPlainText($Property) {
  if (-not $Property) { return '' }

  switch ($Property.type) {
    'title'     { return (($Property.title | ForEach-Object { $_.plain_text }) -join '') }
    'rich_text' { return (($Property.rich_text | ForEach-Object { $_.plain_text }) -join '') }
    'select'    { return $Property.select.name }
    'date'      { return $Property.date.start }
    'checkbox'  { return [string]$Property.checkbox }
    default     { return '' }
  }
}

function Get-WatchCssClass([string]$Color) {
  switch ($Color) {
    'purple_background' { return 'wc-purple' }
    'green_background'  { return 'wc-green' }
    'brown_background'  { return 'wc-brown' }
    'gray_background'   { return 'wc-gray' }
    default             { return 'wc-gray' }
  }
}

function Get-WatchBodyModifier([string]$Text) {
  $lower = $Text.ToLowerInvariant()
  if ($lower.Contains('no fresh signal') -or $lower.Contains('no new updates')) {
    return ' dim'
  }
  return ''
}

function Render-HeadlineItem([string]$Roman, $Record) {
  $sourceMarkup = ''
  if ($Record.Source) {
    $sourceMarkup = '    <span class="src">' + (HtmlEncode $Record.Source) + '</span>'
  }

@"
<div class="hl-item">
  <div class="hl-n">$Roman</div>
  <div>
    <div class="hl-h">$(HtmlEncode $Record.Headline)</div>
    <div class="hl-s">$(HtmlEncode $Record.Summary)</div>
$sourceMarkup
  </div>
</div>
"@
}

function Render-WatchCard($Callout) {
  $title = Get-PlainFromRichText $Callout.callout.rich_text
  $body = ''
  if ($Callout.has_children) {
    $body = Clean-Separator (((Get-NotionChildren $Callout.id) | ForEach-Object { Get-RecursiveBlockText $_ }) -join ' ')
  }
  if (-not $body) {
    $body = Get-RecursiveBlockText $Callout
  }
  if (-not $body) { $body = 'No fresh signal' }
  if ($title -and $body.StartsWith($title + ' ')) {
    $body = $body.Substring($title.Length).Trim()
  }

  $cssClass = Get-WatchCssClass $Callout.callout.color
  $bodyClass = 'wc-body' + (Get-WatchBodyModifier $body)
  $icon = ''
  if ($Callout.callout.icon.type -eq 'emoji') {
    $icon = $Callout.callout.icon.emoji
  }

@"
<div class="watch-card $cssClass">
  <div class="wc-name"><span style="margin-right:.3rem">$(HtmlEncode $icon)</span>$(HtmlEncode $title)</div>
  <div class="$bodyClass">$(HtmlEncode $body)</div>
</div>
"@
}

function Render-PotionCard([string]$Text) {
@"
<div class="potion-card">
  <div class="potion-ico">&#x1F9EA;</div>
  <div class="potion-text">$(HtmlEncode $Text)</div>
</div>
"@
}

function Render-SpellItem($Block) {
  $record = Split-RichTextRecord $Block.to_do.rich_text
  $textParts = @()
  if ($record.Headline) { $textParts += $record.Headline }
  if ($record.Summary) { $textParts += $record.Summary }
  $spellText = Clean-Separator ($textParts -join (' ' + $EmDash + ' '))

  $sourceMarkup = ''
  if ($record.Source) {
    $sourceMarkup = ' <span class="src">' + (HtmlEncode $record.Source) + '</span>'
  }

@"
<li class="spell">
  <div class="spell-box"></div>
  <div class="spell-txt">$(HtmlEncode $spellText)$sourceMarkup</div>
</li>
"@
}

function Render-Template($Slots) {
  $template = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>The Daily Prophet &mdash; {{DATE_DISPLAY}} &middot; {{EDITION}}</title>
<link href="https://fonts.googleapis.com/css2?family=UnifrakturMaguntia&family=IM+Fell+English:ital@0;1&family=IM+Fell+English+SC&family=Cinzel:wght@400;600;900&family=Playfair+Display:ital,wght@0,400;0,700;1,400&display=swap" rel="stylesheet">
<style>
:root{--p:#f2e8d0;--p2:#ecdfc5;--p3:#d9c9a0;--ink:#1a1208;--ink2:#3d2f1a;--ink3:#6b5535;--gold:#b8902a;--gold3:#e8cc7a;--rule:#8b6914;--red:#8b1a1a;--red2:#5c0f0f;--sh:rgba(26,18,8,.32);}
*{margin:0;padding:0;box-sizing:border-box;}
body{background:#18110a;min-height:100vh;padding:2rem 1rem 5rem;font-family:'IM Fell English',Georgia,serif;color:var(--ink);}
.mote{position:fixed;width:2px;height:2px;border-radius:50%;background:var(--gold3);opacity:0;pointer-events:none;animation:mote-rise linear infinite;}
@keyframes mote-rise{0%{transform:translateY(105vh) translateX(0);opacity:0}8%{opacity:.45}92%{opacity:.12}100%{transform:translateY(-30px) translateX(var(--dx,25px));opacity:0}}
.paper{max-width:920px;margin:0 auto;background:var(--p);background-image:linear-gradient(158deg,#f6eed9 0%,#f2e8d0 40%,#ecdfc5 75%,#e4d5b5 100%);box-shadow:0 0 0 1px var(--p3),0 6px 14px var(--sh),0 32px 90px rgba(0,0,0,.7);position:relative;overflow:hidden;animation:paper-unfurl .65s cubic-bezier(.2,.8,.3,1) both;}
@keyframes paper-unfurl{from{opacity:0;transform:translateY(14px) scale(.987)}to{opacity:1;transform:translateY(0) scale(1)}}
.paper::before{content:'';position:absolute;inset:0;pointer-events:none;z-index:1;background:radial-gradient(ellipse at 0% 0%,rgba(90,60,15,.13) 0%,transparent 36%),radial-gradient(ellipse at 100% 100%,rgba(90,60,15,.13) 0%,transparent 36%);}
.masthead{padding:1.8rem 2.5rem 0;text-align:center;border-bottom:3px double var(--rule);position:relative;z-index:2;}
.mh-bar{display:flex;align-items:center;gap:.8rem;margin-bottom:.55rem;}
.mh-rule{flex:1;height:1px;background:linear-gradient(90deg,transparent,var(--rule),transparent);}
.mh-eyebrow{font-family:'IM Fell English SC',serif;font-size:.6rem;letter-spacing:.3em;color:var(--ink3);text-transform:uppercase;white-space:nowrap;}
.mh-title{font-family:'UnifrakturMaguntia',cursive;font-size:clamp(3rem,9vw,6.2rem);color:var(--ink);line-height:.93;text-shadow:2px 2px 0 rgba(180,140,60,.15);margin-bottom:.22rem;}
.mh-subtitle{font-family:'IM Fell English',serif;font-style:italic;font-size:.82rem;color:var(--ink3);letter-spacing:.07em;margin-bottom:.65rem;}
.mh-meta{display:flex;justify-content:space-between;align-items:center;font-family:'IM Fell English SC',serif;font-size:.6rem;letter-spacing:.1em;color:var(--ink2);padding:.42rem 0 .6rem;border-top:1px solid var(--p3);margin-top:.35rem;}
.seal{width:38px;height:38px;border-radius:50%;background:var(--red);border:2px solid var(--red2);display:flex;align-items:center;justify-content:center;font-size:1.25rem;flex-shrink:0;box-shadow:0 0 0 3px var(--p),0 0 0 4px var(--red2);animation:seal-pulse 4s ease-in-out infinite;}
@keyframes seal-pulse{0%,100%{box-shadow:0 0 0 3px var(--p),0 0 0 4px var(--red2)}50%{box-shadow:0 0 0 3px var(--p),0 0 0 4px var(--red2),0 0 16px rgba(139,26,26,.45)}}
.ribbon{background:var(--ink);color:var(--gold3);font-family:'Cinzel',serif;font-size:.66rem;font-weight:600;letter-spacing:.28em;text-transform:uppercase;text-align:center;padding:.33rem 1rem;}
.body{padding:0 2.5rem 2.5rem;position:relative;z-index:2;}
.dispatch{margin:1.3rem 0 0;background:linear-gradient(135deg,rgba(184,144,42,.07),rgba(184,144,42,.03));border:1px solid var(--p3);border-left:4px solid var(--gold);padding:.9rem 1.2rem .85rem 3rem;position:relative;}
.dispatch-owl{position:absolute;left:.65rem;top:50%;transform:translateY(-50%);font-size:1.7rem;animation:owl-bob 3.5s ease-in-out infinite;}
@keyframes owl-bob{0%,100%{transform:translateY(-50%) rotate(-4deg)}50%{transform:translateY(calc(-50% - 5px)) rotate(4deg)}}
.dispatch-head{font-family:'Cinzel',serif;font-size:.7rem;font-weight:600;letter-spacing:.22em;color:var(--ink);text-transform:uppercase;margin-bottom:.15rem;}
.dispatch-sub{font-style:italic;font-size:.8rem;color:var(--ink3);}
.nav-strip{background:rgba(139,105,20,.08);border:1px solid var(--p3);border-top:none;padding:.38rem 1.2rem;font-family:'IM Fell English SC',serif;font-size:.63rem;letter-spacing:.1em;color:var(--ink3);text-align:center;margin-bottom:1rem;}
.sec{display:flex;align-items:center;gap:.65rem;margin:1.5rem 0 .7rem;}
.sec-rl{flex:0 0 16px;height:2px;background:var(--rule);}
.sec-rr{flex:1;height:1px;background:linear-gradient(90deg,var(--rule),transparent);}
.sec-lbl{font-family:'Cinzel',serif;font-size:.66rem;font-weight:600;letter-spacing:.22em;text-transform:uppercase;color:var(--ink2);white-space:nowrap;}
.sec-ico{font-size:.9rem;flex-shrink:0;}
.dbl{border:none;border-top:3px double var(--rule);margin:1.5rem 0;opacity:.5;}
.lead{border:1px solid var(--p3);border-top:3px solid var(--gold);padding:1rem 1.3rem 1rem 2.4rem;margin-bottom:1rem;background:linear-gradient(160deg,rgba(184,144,42,.06),transparent);position:relative;}
.lead-n{position:absolute;top:.9rem;left:.8rem;font-family:'Cinzel',serif;font-size:.65rem;font-weight:900;color:var(--gold);}
.lead-hl{font-family:'Playfair Display',serif;font-weight:700;font-size:1.2rem;color:var(--ink);line-height:1.2;margin-bottom:.32rem;}
.lead-sum{font-size:.88rem;color:var(--ink2);line-height:1.65;margin-bottom:.28rem;}
.src{font-family:'IM Fell English SC',serif;font-size:.63rem;color:var(--gold);letter-spacing:.06em;border-bottom:1px solid rgba(184,144,42,.35);display:inline;}
code{font-family:'Courier New',monospace;font-size:.85rem;background:rgba(139,105,20,.12);padding:.05rem .3rem;}
.hl-cols{display:grid;grid-template-columns:1fr 1fr;gap:0 1.4rem;}
.hl-item{padding:.62rem 0;border-bottom:1px dotted rgba(139,105,20,.3);display:grid;grid-template-columns:auto 1fr;gap:.5rem .75rem;}
.hl-item:last-child{border-bottom:none;}
.hl-n{font-family:'Cinzel',serif;font-size:.63rem;font-weight:900;color:var(--gold);padding-top:.1rem;min-width:1.1rem;}
.hl-h{font-family:'Playfair Display',serif;font-weight:700;font-size:.92rem;color:var(--ink);line-height:1.22;margin-bottom:.16rem;}
.hl-s{font-size:.81rem;color:var(--ink2);line-height:1.55;margin-bottom:.2rem;}
.watch-grid{display:grid;grid-template-columns:1fr 1fr;gap:.85rem;}
.watch-card{border:1px solid var(--p3);padding:.72rem .95rem;position:relative;overflow:hidden;}
.watch-card::after{content:'';position:absolute;bottom:0;left:0;right:0;height:2px;background:linear-gradient(90deg,var(--wc,var(--gold)),transparent);opacity:.4;}
.wc-purple{--wc:#7a5a9a;background:rgba(74,42,122,.04);}
.wc-green{--wc:#2a6b2a;background:rgba(42,107,42,.04);}
.wc-brown{--wc:#8b5a2a;background:rgba(92,58,26,.04);}
.wc-gray{--wc:#888780;background:rgba(136,135,128,.04);}
.wc-name{font-family:'Cinzel',serif;font-size:.65rem;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--ink);margin-bottom:.28rem;}
.wc-body{font-size:.81rem;color:var(--ink2);line-height:1.58;}
.wc-body.dim{color:var(--ink3);font-style:italic;}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:1.8rem;}
.potion-card{background:rgba(42,107,42,.04);border:1px solid rgba(42,107,42,.2);border-left:3px solid #2a6b2a;padding:.72rem 1rem .72rem 2.7rem;position:relative;margin-bottom:.45rem;}
.potion-ico{position:absolute;left:.7rem;top:.62rem;font-size:.95rem;}
.potion-text{font-size:.84rem;color:var(--ink2);line-height:1.62;}
.spell-list{list-style:none;}
.spell{display:flex;align-items:flex-start;gap:.62rem;padding:.42rem 0;border-bottom:1px dotted rgba(139,105,20,.25);}
.spell:last-child{border-bottom:none;}
.spell-box{width:13px;height:13px;border:1.5px solid var(--gold);flex-shrink:0;margin-top:.22rem;background:rgba(184,144,42,.05);}
.spell-txt{font-size:.84rem;color:var(--ink2);line-height:1.55;}
.map-note{border-left:3px solid var(--p3);padding:.6rem 1rem;font-size:.84rem;color:var(--ink3);font-style:italic;line-height:1.62;background:rgba(139,105,20,.03);}
.tl-wrap{border:1px solid var(--p3);overflow:hidden;margin-top:.5rem;}
.tl-summary{display:flex;align-items:center;gap:.55rem;padding:.5rem .85rem;cursor:pointer;font-family:'Cinzel',serif;font-size:.63rem;font-weight:600;letter-spacing:.15em;text-transform:uppercase;color:var(--ink3);background:rgba(139,105,20,.04);user-select:none;list-style:none;}
.tl-summary::-webkit-details-marker{display:none;}
.tl-body{padding:.65rem .85rem;font-size:.81rem;color:var(--ink3);font-style:italic;border-top:1px dotted var(--p3);}
.footer{text-align:center;padding:1.1rem 2.5rem 1.7rem;border-top:3px double var(--rule);}
.footer-txt{font-family:'IM Fell English SC',serif;font-size:.6rem;letter-spacing:.18em;color:var(--ink3);}
.footer-motto{font-family:'IM Fell English',serif;font-style:italic;font-size:.79rem;color:var(--ink2);margin-top:.28rem;}
.fi{animation:fi .45s ease-out both;}
.d1{animation-delay:.08s}.d2{animation-delay:.16s}.d3{animation-delay:.24s}.d4{animation-delay:.32s}.d5{animation-delay:.4s}.d6{animation-delay:.48s}
@keyframes fi{from{opacity:0;transform:translateY(5px)}to{opacity:1;transform:translateY(0)}}
@media(max-width:600px){.masthead,.body,.footer{padding-left:1.4rem;padding-right:1.4rem;}.two-col,.hl-cols,.watch-grid{grid-template-columns:1fr;}}
</style>
</head>
<body>
<div class="mote" style="left:7%;animation-duration:21s;animation-delay:0s;--dx:20px"></div>
<div class="mote" style="left:22%;animation-duration:27s;animation-delay:5s;--dx:-18px"></div>
<div class="mote" style="left:41%;animation-duration:18s;animation-delay:9s;--dx:30px"></div>
<div class="mote" style="left:60%;animation-duration:24s;animation-delay:2s;--dx:-25px"></div>
<div class="mote" style="left:78%;animation-duration:20s;animation-delay:7s;--dx:22px"></div>
<div class="mote" style="left:91%;animation-duration:29s;animation-delay:12s;--dx:-20px"></div>

<div class="paper">
  <div class="masthead">
    <div class="mh-bar">
      <div class="mh-rule"></div>
      <div class="mh-eyebrow">Curated from your universe &middot; Est. by order of the Ministry</div>
      <div class="mh-rule"></div>
    </div>
    <div class="mh-title">The Daily Prophet</div>
    <div class="mh-subtitle">Autonomous Intelligence Dispatch &middot; America / New York</div>
    <div class="mh-meta">
      <span>Vol. I &nbsp;&middot;&nbsp; Issue&nbsp;<strong>{{ISSUE_ID}}</strong></span>
      <div class="seal">&#x1F989;</div>
      <span>{{DATE_DISPLAY}} &nbsp;&middot;&nbsp; Owl-posted at dawn</span>
    </div>
  </div>

  <div class="ribbon">{{RIBBON_ICON}}&nbsp;&nbsp;{{EDITION}}&nbsp;&nbsp;{{RIBBON_ICON}}</div>

  <div class="body">
    <div class="dispatch fi d1">
      <div class="dispatch-owl">&#x1F989;</div>
      <div class="dispatch-head">The Daily Prophet</div>
      <div class="dispatch-sub">{{DISPATCH_SUBTITLE}}</div>
    </div>
    <div class="nav-strip fi d1">{{NAV_STRIP}}</div>

    <hr class="dbl">

    <div class="sec fi d2">
      <div class="sec-rl"></div><span class="sec-ico">&#x1F5DE;&#xFE0F;</span>
      <span class="sec-lbl">Front Page &mdash; Top Headlines</span>
      <div class="sec-rr"></div>
    </div>

    <div class="lead fi d2">
      <div class="lead-n">I.</div>
      <div class="lead-hl">{{LEAD_HEADLINE}}</div>
      <div class="lead-sum">{{LEAD_SUMMARY}}</div>
      <span class="src">{{LEAD_SOURCE}}</span>
    </div>

    <div class="hl-cols fi d3">
      <div>{{HEADLINES_COL_LEFT}}</div>
      <div>{{HEADLINES_COL_RIGHT}}</div>
    </div>

    <hr class="dbl">

    <div class="sec fi d3">
      <div class="sec-rl"></div><span class="sec-ico">&#x1F3F0;</span>
      <span class="sec-lbl">Hogwarts Watchlist &mdash; Your Active Worlds</span>
      <div class="sec-rr"></div>
    </div>
    <div class="watch-grid fi d4">{{WATCHLIST_CARDS}}</div>

    <hr class="dbl">

    <div class="two-col fi d5">
      <div>
        <div class="sec">
          <div class="sec-rl"></div><span class="sec-ico">&#x1F9EA;</span>
          <span class="sec-lbl">Potion Notes</span>
          <div class="sec-rr"></div>
        </div>
        {{POTION_NOTES}}

        <div class="sec" style="margin-top:1.3rem">
          <div class="sec-rl"></div><span class="sec-ico">&#x1FA84;</span>
          <span class="sec-lbl">Today's Spells</span>
          <div class="sec-rr"></div>
        </div>
        <ul class="spell-list">{{SPELLS}}</ul>
      </div>
      <div>
        <div class="sec">
          <div class="sec-rl"></div><span class="sec-ico">&#x1F5FA;&#xFE0F;</span>
          <span class="sec-lbl">The Map</span>
          <div class="sec-rr"></div>
        </div>
        <div class="map-note">{{THE_MAP}}</div>

        <div class="sec" style="margin-top:1.3rem">
          <div class="sec-rl"></div><span class="sec-ico">&#x1F30D;</span>
          <span class="sec-lbl">Translation</span>
          <div class="sec-rr"></div>
        </div>
        <details class="tl-wrap">
          <summary class="tl-summary">&#x1F30D; &nbsp;Only when enabled</summary>
          <div class="tl-body">
            <div style="margin-bottom:.28rem">Target language: {{TARGET_LANGUAGE}}</div>
            <div>Translated briefing: {{TRANSLATED_BRIEFING}}</div>
          </div>
        </details>
      </div>
    </div>
  </div>

  <div class="footer fi d6">
    <div class="footer-txt">Printed by Order of the Ministry of Magic &nbsp;&middot;&nbsp; {{EDITION}} &nbsp;&middot;&nbsp; Issue {{ISSUE_ID}} &nbsp;&middot;&nbsp; {{DATE_DISPLAY}}</div>
    <a href="{{ARCHIVE_HREF}}" style="font-family:'IM Fell English SC',serif;font-size:.6rem;letter-spacing:.18em;color:var(--gold);text-decoration:none;">&mdash; The Archive &mdash;</a>
    <div class="footer-motto">"Owl-posted at dawn. Curated from your universe."</div>
  </div>
</div>
</body>
</html>
'@

  $html = $template
  foreach ($key in $Slots.Keys) {
    $html = $html.Replace("{{$key}}", [string]$Slots[$key])
  }
  return $html
}

function Render-NoIssuePage([string]$DateDisplay) {
  $slots = @{
    DATE_DISPLAY        = HtmlEncode $DateDisplay
    EDITION             = 'Morning Edition'
    RIBBON_ICON         = '&#9728;'
    ISSUE_ID            = '&mdash;'
    DISPATCH_SUBTITLE   = 'No issue yet - check back after 6 AM ET'
    NAV_STRIP           = 'Front Page &middot; Hogwarts Watchlist &middot; Potion Notes &middot; Today''s Spells &middot; The Map'
    LEAD_HEADLINE       = 'No issue yet'
    LEAD_SUMMARY        = 'Check back after 6 AM ET for the next owl-posted morning edition.'
    LEAD_SOURCE         = ''
    HEADLINES_COL_LEFT  = ''
    HEADLINES_COL_RIGHT = ''
    WATCHLIST_CARDS     = ''
    POTION_NOTES        = (Render-PotionCard 'No potion notes have been filed for this edition yet.')
    SPELLS              = '<li class="spell"><div class="spell-box"></div><div class="spell-txt">Wait for the morning issue to arrive.</div></li>'
    THE_MAP             = 'The newsroom is still gathering its morning dispatch.'
    TARGET_LANGUAGE     = '(not set)'
    TRANSLATED_BRIEFING = '(disabled for this issue)'
    ARCHIVE_HREF        = 'issues/'
  }

  return Render-Template $slots
}

function Get-IssueArchiveFileName([string]$IssueId) {
  $safeName = $IssueId
  foreach ($invalid in [System.IO.Path]::GetInvalidFileNameChars()) {
    $safeName = $safeName.Replace([string]$invalid, '-')
  }
  return $safeName + '.html'
}

function Get-IssueRenderData($Page, [string]$ArchiveHref) {
  $props = $Page.properties
  $blocks = @(Get-NotionChildren $Page.id)

  $callouts = @($blocks | Where-Object { $_.type -eq 'callout' })
  $dispatchCallout = $callouts[0]
  $navCallout = $callouts[1]

  $dispatchSubtitle = ''
  if ($dispatchCallout.has_children) {
    $dispatchSubtitle = Clean-Separator (((Get-NotionChildren $dispatchCallout.id) | ForEach-Object { Get-RecursiveBlockText $_ }) -join ' ')
  }
  if (-not $dispatchSubtitle) {
    $dispatchSubtitle = Get-PlainFromRichText $dispatchCallout.callout.rich_text
  }

  $navStrip = Get-PlainFromRichText $navCallout.callout.rich_text
  if (-not $navStrip) {
    $navStrip = Get-RecursiveBlockText $navCallout
  }

  $headings = @($blocks | Where-Object { $_.type -eq 'heading_1' })
  $frontHeading  = $headings[0]
  $watchHeading  = $headings[1]
  $potionHeading = $headings[2]
  $spellsHeading = $headings[3]
  $mapHeading    = $headings[4]

  $frontIndex  = [array]::IndexOf($blocks, $frontHeading)
  $watchIndex  = [array]::IndexOf($blocks, $watchHeading)
  $potionIndex = [array]::IndexOf($blocks, $potionHeading)
  $spellsIndex = [array]::IndexOf($blocks, $spellsHeading)
  $mapIndex    = [array]::IndexOf($blocks, $mapHeading)

  $leadBlock = $blocks[$frontIndex + 1]
  $leadRecord = Split-RichTextRecord $leadBlock.quote.rich_text

  $frontColumnList = $blocks[$frontIndex + 2]
  $frontColumns = Get-ColumnChildren $frontColumnList.id

  $romanNumerals = @('II.','III.','IV.','V.','VI.','VII.')
  $romanIndex = 0

  $leftHeadlineMarkup = @()
  foreach ($item in @($frontColumns[0])) {
    $leftHeadlineMarkup += Render-HeadlineItem -Roman $romanNumerals[$romanIndex] -Record (Split-RichTextRecord $item.bulleted_list_item.rich_text)
    $romanIndex++
  }

  $rightHeadlineMarkup = @()
  foreach ($item in @($frontColumns[1])) {
    $rightHeadlineMarkup += Render-HeadlineItem -Roman $romanNumerals[$romanIndex] -Record (Split-RichTextRecord $item.bulleted_list_item.rich_text)
    $romanIndex++
  }

  $watchColumnList = $blocks[$watchIndex + 1]
  $watchColumns = Get-ColumnChildren $watchColumnList.id
  $watchCards = @()
  foreach ($column in @($watchColumns)) {
    foreach ($callout in @($column)) {
      $watchCards += Render-WatchCard $callout
    }
  }

  $potionCallout = $blocks[$potionIndex + 1]
  $potionText = Get-RecursiveBlockText $potionCallout
  if (-not $potionText) {
    $potionText = Get-PlainFromRichText $potionCallout.callout.rich_text
  }
  $potionMarkup = Render-PotionCard $potionText

  $spellBlocks = @()
  for ($i = $spellsIndex + 1; $i -lt $mapIndex; $i++) {
    if ($blocks[$i].type -eq 'to_do') {
      $spellBlocks += $blocks[$i]
    }
  }

  $spellMarkup = @()
  foreach ($spellBlock in $spellBlocks) {
    $spellMarkup += Render-SpellItem $spellBlock
  }

  $mapBlock = $blocks[$mapIndex + 1]
  $mapText = Get-PlainFromRichText $mapBlock.quote.rich_text

  $editionRaw = Get-PropertyPlainText $props.Edition
  $editionLabel = Get-EditionLabel $editionRaw
  $issueId = Get-PropertyPlainText $props.'Issue ID'
  $issueDate = Get-PropertyPlainText $props.Date
  $dateDisplay = Get-DateDisplay $issueDate
  $translationEnabled = $false
  if ($props.PSObject.Properties.Name -contains 'Translation') {
    $translationEnabled = [bool]$props.Translation.checkbox
  }

  $targetLanguage = ''
  if ($props.PSObject.Properties.Name -contains 'Target language') {
    $targetLanguage = Get-PropertyPlainText $props.'Target language'
  }
  if (-not $targetLanguage) {
    $targetLanguage = '(not set)'
  }

  $translatedBriefing = '(disabled for this issue)'
  if ($translationEnabled) {
    if ($props.PSObject.Properties.Name -contains 'Translated briefing') {
      $translatedBriefing = Get-PropertyPlainText $props.'Translated briefing'
    }
    if (-not $translatedBriefing) {
      $translatedBriefing = '(not set)'
    }
  }

  $slots = @{
    DATE_DISPLAY        = HtmlEncode $dateDisplay
    EDITION             = HtmlEncode $editionLabel
    RIBBON_ICON         = Get-RibbonIcon $editionLabel
    ISSUE_ID            = HtmlEncode $issueId
    DISPATCH_SUBTITLE   = HtmlEncode $dispatchSubtitle
    NAV_STRIP           = HtmlEncode $navStrip
    LEAD_HEADLINE       = HtmlEncode $leadRecord.Headline
    LEAD_SUMMARY        = HtmlEncode $leadRecord.Summary
    LEAD_SOURCE         = HtmlEncode $leadRecord.Source
    HEADLINES_COL_LEFT  = ($leftHeadlineMarkup -join "`n")
    HEADLINES_COL_RIGHT = ($rightHeadlineMarkup -join "`n")
    WATCHLIST_CARDS     = ($watchCards -join "`n")
    POTION_NOTES        = $potionMarkup
    SPELLS              = ($spellMarkup -join "`n")
    THE_MAP             = HtmlEncode $mapText
    TARGET_LANGUAGE     = HtmlEncode $targetLanguage
    TRANSLATED_BRIEFING = HtmlEncode $translatedBriefing
    ARCHIVE_HREF        = $ArchiveHref
  }

  return [pscustomobject]@{
    Html      = Render-Template $slots
    IssueId   = $issueId
    IssueDate = $issueDate
    Edition   = $editionLabel
  }
}

function Get-ArchiveBadgeClass([string]$EditionLabel) {
  switch ($EditionLabel) {
    'Morning Edition' { return 'badge-morning' }
    'Evening Edition' { return 'badge-evening' }
    'Breaking News'   { return 'badge-breaking' }
    default           { return 'badge-morning' }
  }
}

function Render-ArchiveIndex($Pages) {
  $rows = @()
  foreach ($page in @($Pages)) {
    $props = $page.properties
    $issueId = Get-PropertyPlainText $props.'Issue ID'
    if (-not $issueId) { continue }

    $issueDate = Get-PropertyPlainText $props.Date
    $editionLabel = Get-EditionLabel (Get-PropertyPlainText $props.Edition)
    $badgeClass = Get-ArchiveBadgeClass $editionLabel
    $fileName = Get-IssueArchiveFileName $issueId
    $dateDisplay = Get-DateDisplay $issueDate

    $rows += @"
<div class="archive-row">
  <div class="archive-date">$(HtmlEncode $dateDisplay)</div>
  <div><span class="badge $badgeClass">$(HtmlEncode $editionLabel)</span></div>
  <div class="archive-issue">$(HtmlEncode $issueId)</div>
  <div><a class="archive-link" href="$(HtmlEncode $fileName)">issues/$(HtmlEncode $fileName)</a></div>
</div>
"@
  }

  $archiveRows = if ($rows.Count -gt 0) { $rows -join "`r`n" } else { '<div class="archive-empty">No archived issues yet.</div>' }

  return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>The Daily Prophet &mdash; Archive</title>
<link href="https://fonts.googleapis.com/css2?family=UnifrakturMaguntia&family=IM+Fell+English:ital@0;1&family=IM+Fell+English+SC&family=Cinzel:wght@400;600;900&family=Playfair+Display:ital,wght@0,400;0,700;1,400&display=swap" rel="stylesheet">
<style>
:root{--p:#f2e8d0;--p2:#ecdfc5;--p3:#d9c9a0;--ink:#1a1208;--ink2:#3d2f1a;--ink3:#6b5535;--gold:#b8902a;--gold3:#e8cc7a;--rule:#8b6914;--red:#8b1a1a;--red2:#5c0f0f;--sh:rgba(26,18,8,.32);}
*{margin:0;padding:0;box-sizing:border-box;}
body{background:#18110a;min-height:100vh;padding:2rem 1rem 5rem;font-family:'IM Fell English',Georgia,serif;color:var(--ink);}
.paper{max-width:920px;margin:0 auto;background:var(--p);background-image:linear-gradient(158deg,#f6eed9 0%,#f2e8d0 40%,#ecdfc5 75%,#e4d5b5 100%);box-shadow:0 0 0 1px var(--p3),0 6px 14px var(--sh),0 32px 90px rgba(0,0,0,.7);position:relative;overflow:hidden;}
.paper::before{content:'';position:absolute;inset:0;pointer-events:none;z-index:1;background:radial-gradient(ellipse at 0% 0%,rgba(90,60,15,.13) 0%,transparent 36%),radial-gradient(ellipse at 100% 100%,rgba(90,60,15,.13) 0%,transparent 36%);}
.masthead{padding:1.8rem 2.5rem 0;text-align:center;border-bottom:3px double var(--rule);position:relative;z-index:2;}
.mh-bar{display:flex;align-items:center;gap:.8rem;margin-bottom:.55rem;}
.mh-rule{flex:1;height:1px;background:linear-gradient(90deg,transparent,var(--rule),transparent);}
.mh-eyebrow{font-family:'IM Fell English SC',serif;font-size:.6rem;letter-spacing:.3em;color:var(--ink3);text-transform:uppercase;white-space:nowrap;}
.mh-title{font-family:'UnifrakturMaguntia',cursive;font-size:clamp(3rem,9vw,6.2rem);color:var(--ink);line-height:.93;text-shadow:2px 2px 0 rgba(180,140,60,.15);margin-bottom:.22rem;}
.mh-subtitle{font-family:'IM Fell English',serif;font-style:italic;font-size:.82rem;color:var(--ink3);letter-spacing:.07em;margin-bottom:.65rem;}
.mh-meta{display:flex;justify-content:center;align-items:center;font-family:'IM Fell English SC',serif;font-size:.6rem;letter-spacing:.1em;color:var(--ink2);padding:.42rem 0 .6rem;border-top:1px solid var(--p3);margin-top:.35rem;}
.body{padding:1.4rem 2.5rem 2.5rem;position:relative;z-index:2;}
.archive-head{font-family:'Cinzel',serif;font-size:.7rem;font-weight:600;letter-spacing:.22em;text-transform:uppercase;color:var(--ink2);margin-bottom:.8rem;}
.archive-list{border:1px solid var(--p3);background:rgba(255,255,255,.12);}
.archive-row{display:grid;grid-template-columns:2.3fr 1.4fr 1.4fr 2.4fr;gap:.8rem;align-items:center;padding:.75rem 1rem;border-bottom:1px dotted rgba(139,105,20,.28);}
.archive-row:last-child{border-bottom:none;}
.archive-date,.archive-issue{font-size:.88rem;color:var(--ink2);}
.archive-link{font-family:'IM Fell English SC',serif;font-size:.68rem;letter-spacing:.06em;color:var(--gold);text-decoration:none;}
.archive-link:hover{text-decoration:underline;}
.badge{display:inline-block;padding:.22rem .6rem;border:1px solid currentColor;font-family:'Cinzel',serif;font-size:.62rem;font-weight:600;letter-spacing:.14em;text-transform:uppercase;}
.badge-morning{color:var(--gold);background:rgba(184,144,42,.08);}
.badge-evening{color:#6b4ea0;background:rgba(107,78,160,.08);}
.badge-breaking{color:var(--red);background:rgba(139,26,26,.08);}
.archive-empty{padding:1rem;font-size:.9rem;color:var(--ink3);font-style:italic;}
.footer{text-align:center;padding:1.1rem 2.5rem 1.7rem;border-top:3px double var(--rule);}
.footer-txt{font-family:'IM Fell English SC',serif;font-size:.6rem;letter-spacing:.18em;color:var(--ink3);}
.footer-link{display:inline-block;margin-top:.45rem;font-family:'IM Fell English SC',serif;font-size:.6rem;letter-spacing:.18em;color:var(--gold);text-decoration:none;}
@media(max-width:700px){.masthead,.body,.footer{padding-left:1.4rem;padding-right:1.4rem;}.archive-row{grid-template-columns:1fr;gap:.35rem;}}
</style>
</head>
<body>
<div class="paper">
  <div class="masthead">
    <div class="mh-bar">
      <div class="mh-rule"></div>
      <div class="mh-eyebrow">Curated from your universe &middot; Est. by order of the Ministry</div>
      <div class="mh-rule"></div>
    </div>
    <div class="mh-title">The Daily Prophet &mdash; Archive</div>
    <div class="mh-subtitle">Past owl-posted editions from the Daily Prophet Issues desk</div>
    <div class="mh-meta">Collected from DB | Daily Prophet Issues</div>
  </div>
  <div class="body">
    <div class="archive-head">Issue Ledger</div>
    <div class="archive-list">
$archiveRows
    </div>
  </div>
  <div class="footer">
    <div class="footer-txt">Printed by Order of the Ministry of Magic &nbsp;&middot;&nbsp; Archive Ledger</div>
    <a class="footer-link" href="../">Return to Today&apos;s Issue</a>
  </div>
</div>
</body>
</html>
"@
}

$issueDate = Get-IssueDate
$dateDisplay = Get-DateDisplay $issueDate

$queryBody = @{
  filter = @{
    and = @(
      @{ property = 'Date'; date = @{ equals = $issueDate } },
      @{ property = 'Edition'; select = @{ equals = 'Morning' } }
    )
  }
}

$queryResponse = Invoke-DatabaseQuery $queryBody
$allIssuePages = Get-AllIssuePages

if (-not (Test-Path -LiteralPath $IssuesDirectory)) {
  New-Item -ItemType Directory -Path $IssuesDirectory | Out-Null
}

if (-not $queryResponse.results -or $queryResponse.results.Count -eq 0) {
  $html = Render-NoIssuePage $dateDisplay
  Write-Utf8File -Path $OutputPath -Content $html
  Write-Utf8File -Path $ArchiveIndexPath -Content (Render-ArchiveIndex $allIssuePages)
  Write-Host ('Written: index.html ' + $EmDash + ' Issue ' + $EmDash + ' ' + $EmDash + ' ' + $issueDate)
  exit 0
}

$currentPage = $queryResponse.results[0]
$currentRender = Get-IssueRenderData -Page $currentPage -ArchiveHref 'issues/'
Write-Utf8File -Path $OutputPath -Content $currentRender.Html

$archiveCount = 0
foreach ($issuePage in @($allIssuePages)) {
  $issueId = Get-PropertyPlainText $issuePage.properties.'Issue ID'
  if (-not $issueId) { continue }

  $renderedIssue = if ($issuePage.id -eq $currentPage.id) {
    Get-IssueRenderData -Page $issuePage -ArchiveHref './'
  } else {
    Get-IssueRenderData -Page $issuePage -ArchiveHref './'
  }

  $archivePath = Join-Path $IssuesDirectory (Get-IssueArchiveFileName $renderedIssue.IssueId)
  Write-Utf8File -Path $archivePath -Content $renderedIssue.Html
  $archiveCount++
}

Write-Utf8File -Path $ArchiveIndexPath -Content (Render-ArchiveIndex $allIssuePages)
Write-Host ('Written: index.html ' + $EmDash + ' Issue ' + $currentRender.IssueId + ' ' + $EmDash + ' ' + $issueDate)
Write-Host ('Archived: ' + $archiveCount + ' issues')
