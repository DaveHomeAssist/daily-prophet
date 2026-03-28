[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$DatabaseId = '331255fc8f4480d59f92ffb703398031'
$OutputPath = Join-Path $PSScriptRoot 'index.html'

function Get-EnvValue([string]$Name) {
  $processValue = [Environment]::GetEnvironmentVariable($Name, 'Process')
  if ($processValue) { return $processValue }
  return [Environment]::GetEnvironmentVariable($Name, 'User')
}

function Get-IssueDate() {
  $override = Get-EnvValue 'ISSUE_DATE'
  if ($override) { return $override }
  return [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'Eastern Standard Time').ToString('yyyy-MM-dd')
}

function Get-NotionHeaders() {
  $token = Get-EnvValue 'NOTION_API_KEY'
  if (-not $token) {
    throw 'NOTION_API_KEY is missing from both the current process and User environment variables.'
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
    $params.Body = ($Body | ConvertTo-Json -Depth 20)
  }

  return Invoke-RestMethod @params
}

function Get-NotionChildren([string]$BlockId) {
  $items = @()
  $cursor = $null

  do {
    $uri = "https://api.notion.com/v1/blocks/$($BlockId -replace '-','')/children?page_size=100"
    if ($cursor) {
      $uri += "&start_cursor=$cursor"
    }

    $resp = Invoke-NotionJson -Method Get -Uri $uri
    $items += $resp.results
    $cursor = $resp.next_cursor
  } while ($cursor)

  return $items
}

function Get-PropertyPlainText($Prop) {
  if (-not $Prop) { return '' }
  switch ($Prop.type) {
    'title'      { return (($Prop.title | ForEach-Object plain_text) -join '') }
    'rich_text'  { return (($Prop.rich_text | ForEach-Object plain_text) -join '') }
    'select'     { return $Prop.select.name }
    'multi_select' { return (($Prop.multi_select | ForEach-Object name) -join ', ') }
    'date'       { return $Prop.date.start }
    'checkbox'   { return [string]$Prop.checkbox }
    default      { return '' }
  }
}

function Get-RichTextPlain($RichText) {
  if (-not $RichText) { return '' }

  $parts = foreach ($item in $RichText) {
    if ($item.type -eq 'mention' -and $item.mention.type -eq 'page' -and $item.plain_text -eq 'Untitled') {
      continue
    }
    $item.plain_text
  }

  $text = ($parts -join '')
  $text = $text -replace '\s*[\p{Pd}]\s*Untitled\b', ''
  $text = $text -replace '\s{2,}', ' '
  $text = $text.Trim()
  $text = $text -replace '\s*[\p{Pd}]+\s*$', ''
  return $text.Trim()
}

function Get-BlockText($Block) {
  $payload = $Block.$($Block.type)
  if (-not $payload) { return '' }
  if ($payload.PSObject.Properties.Name -contains 'rich_text') {
    return Get-RichTextPlain $payload.rich_text
  }
  return ''
}

function Split-Headline([string]$Text) {
  $parts = [regex]::Split($Text, '\s*[\p{Pd}]\s*', 3)
  $headline = if ($parts.Count -gt 0) { $parts[0].Trim() } else { $Text.Trim() }
  $deck = if ($parts.Count -gt 1) { $parts[1].Trim() } else { '' }
  return @{
    Headline = $headline
    Deck     = $deck
  }
}

function Get-ColumnBlocks([string]$ColumnListId) {
  $columns = @()
  foreach ($column in (Get-NotionChildren $ColumnListId)) {
    $columns += ,(Get-NotionChildren $column.id)
  }
  return $columns
}

function Get-WatchlistItem([object]$Callout) {
  $payload = $Callout.callout
  $title = Get-RichTextPlain $payload.rich_text
  $bodyBlocks = @(Get-NotionChildren $Callout.id)
  $body = (($bodyBlocks | ForEach-Object { Get-BlockText $_ }) -join ' ').Trim()
  $color = switch ($payload.color) {
    'purple_background' { 'purple' }
    'green_background'  { 'green' }
    'brown_background'  { 'brown' }
    'gray_background'   { 'gray' }
    default             { 'gray' }
  }

  return @{
    Title = $title
    Body  = $body
    Icon  = $payload.icon.emoji
    Tone  = $color
  }
}

function HtmlEncode([string]$Text) {
  if ($null -eq $Text) { return '' }
  return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Render-FrontPage($FrontLead, $FrontColumns) {
  $items = foreach ($column in $FrontColumns) {
    foreach ($item in $column) {
      $parsed = Split-Headline (Get-BlockText $item)
      @"
<article class="story">
  <h3>$(HtmlEncode $parsed.Headline)</h3>
  <p>$(HtmlEncode $parsed.Deck)</p>
</article>
"@
    }
  }

  return @"
<section class="section">
  <div class="section-kicker">Front Page</div>
  <blockquote class="lead-quote">$(HtmlEncode $FrontLead)</blockquote>
  <div class="columns two">
$(($items -join "`n"))
  </div>
</section>
"@
}

function Render-Watchlist($Items) {
  $cards = foreach ($item in $Items) {
    @"
<article class="watch-card tone-$($item.Tone)">
  <div class="watch-icon">$(HtmlEncode $item.Icon)</div>
  <h3>$(HtmlEncode $item.Title)</h3>
  <p>$(HtmlEncode $item.Body)</p>
</article>
"@
  }

  return @"
<section class="section">
  <div class="section-kicker">Hogwarts Watchlist</div>
  <div class="columns two">
$(($cards -join "`n"))
  </div>
</section>
"@
}

function Render-Spells($Todos) {
  $items = foreach ($todo in $Todos) {
    $text = Get-BlockText $todo
    $checked = if ($todo.to_do.checked) { 'checked' } else { '' }
    @"
<label class="spell">
  <input type="checkbox" disabled $checked>
  <span>$(HtmlEncode $text)</span>
</label>
"@
  }

  return @"
<section class="section">
  <div class="section-kicker">Today's Spells</div>
  <div class="spells">
$(($items -join "`n"))
  </div>
</section>
"@
}

function Render-Page($Data) {
  $dateDisplay = [datetime]::Parse($Data.Date).ToString('dddd, MMMM d, yyyy')
  $sourcesHtml = if ($Data.Sources) { HtmlEncode $Data.Sources } else { 'Notion' }

  return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>The Daily Prophet — $(HtmlEncode $dateDisplay) · $(HtmlEncode $Data.Edition)</title>
<link href="https://fonts.googleapis.com/css2?family=UnifrakturMaguntia&family=IM+Fell+English:ital@0;1&family=IM+Fell+English+SC&family=Cinzel:wght@400;600;900&family=Playfair+Display:ital,wght@0,400;0,700;1,400&display=swap" rel="stylesheet">
<style>
:root {
  --p:#f2e8d0;--p2:#ecdfc5;--p3:#d9c9a0;
  --ink:#1a1208;--ink2:#3d2f1a;--ink3:#6b5535;
  --gold:#b8902a;--gold3:#e8cc7a;--rule:#8b6914;
  --red:#8b1a1a;--red2:#5c0f0f;--sh:rgba(26,18,8,.32);
}
*{margin:0;padding:0;box-sizing:border-box;}
body{background:#18110a;min-height:100vh;padding:2rem 1rem 5rem;font-family:'IM Fell English',Georgia,serif;color:var(--ink);}
.mote{position:fixed;width:2px;height:2px;border-radius:50%;background:var(--gold3);opacity:0;pointer-events:none;animation:mote-rise linear infinite;}
@keyframes mote-rise{0%{transform:translateY(105vh) translateX(0);opacity:0}8%{opacity:.45}92%{opacity:.12}100%{transform:translateY(-30px) translateX(var(--dx,25px));opacity:0}}
.paper{max-width:920px;margin:0 auto;background:var(--p);background-image:linear-gradient(158deg,#f6eed9 0%,#f2e8d0 40%,#ecdfc5 75%,#e4d5b5 100%);box-shadow:0 0 0 1px var(--p3),0 22px 60px var(--sh);position:relative;overflow:hidden}
.paper::before{content:'';position:absolute;inset:0;background:radial-gradient(circle at top left,rgba(255,255,255,.18),transparent 32%),radial-gradient(circle at bottom right,rgba(139,105,20,.08),transparent 30%);pointer-events:none}
.paper-inner{padding:1.5rem 1.2rem 2rem;position:relative}
.nameplate{border-top:3px double var(--rule);border-bottom:3px double var(--rule);padding:1rem 0 .85rem;text-align:center}
.kicker{font-family:'Cinzel',serif;font-size:.72rem;letter-spacing:.2em;text-transform:uppercase;color:var(--ink3);margin-bottom:.35rem}
.flag{font-family:'UnifrakturMaguntia',serif;font-size:clamp(2.8rem,8vw,5.2rem);line-height:1;color:var(--ink)}
.sub{font-family:'IM Fell English SC',serif;font-size:.9rem;letter-spacing:.08em;color:var(--ink2);margin-top:.45rem}
.meta{display:flex;justify-content:space-between;gap:1rem;flex-wrap:wrap;border-bottom:1px solid var(--rule);padding:.7rem 0 .8rem;margin-bottom:1rem;font-family:'Cinzel',serif;font-size:.78rem;letter-spacing:.08em;text-transform:uppercase;color:var(--ink3)}
.summary{font-family:'Playfair Display',serif;font-size:1.08rem;line-height:1.65;color:var(--ink2);padding:.35rem 0 1rem;border-bottom:1px solid rgba(139,105,20,.35)}
.summary strong{font-family:'Cinzel',serif;font-size:.8rem;letter-spacing:.12em;text-transform:uppercase;color:var(--red);margin-right:.55rem}
.section{padding:1.15rem 0;border-bottom:1px solid rgba(139,105,20,.35)}
.section:last-of-type{border-bottom:none}
.section-kicker{font-family:'Cinzel',serif;font-size:.8rem;letter-spacing:.18em;text-transform:uppercase;color:var(--red);margin-bottom:.7rem}
.lead-quote{font-family:'Playfair Display',serif;font-size:1.45rem;line-height:1.35;border-left:4px solid var(--rule);padding:.2rem 0 .2rem .9rem;margin-bottom:1rem}
.columns{display:grid;gap:1rem}
.columns.two{grid-template-columns:repeat(2,minmax(0,1fr))}
.story{background:rgba(255,255,255,.24);border:1px solid rgba(139,105,20,.28);padding:.9rem .95rem}
.story h3{font-family:'Playfair Display',serif;font-size:1.15rem;line-height:1.2;margin-bottom:.35rem}
.story p{line-height:1.55;color:var(--ink2)}
.watch-card{border:1px solid rgba(139,105,20,.28);padding:.95rem;background:rgba(255,255,255,.18)}
.watch-card h3{font-family:'Cinzel',serif;font-size:1rem;letter-spacing:.06em;text-transform:uppercase;margin-bottom:.35rem}
.watch-card p{line-height:1.55;color:var(--ink2)}
.watch-icon{font-size:1.1rem;margin-bottom:.4rem}
.tone-purple{box-shadow:inset 0 0 0 2px rgba(92,15,15,.04);background:linear-gradient(180deg,rgba(255,255,255,.24),rgba(107,85,53,.06))}
.tone-green{background:linear-gradient(180deg,rgba(255,255,255,.24),rgba(184,144,42,.08))}
.tone-brown{background:linear-gradient(180deg,rgba(255,255,255,.24),rgba(61,47,26,.08))}
.tone-gray{background:linear-gradient(180deg,rgba(255,255,255,.24),rgba(107,85,53,.05))}
.note-box{background:rgba(255,255,255,.26);border:1px solid rgba(139,105,20,.28);padding:1rem 1rem 1.05rem;font-size:1.06rem;line-height:1.7}
.spells{display:grid;gap:.7rem}
.spell{display:grid;grid-template-columns:20px 1fr;gap:.7rem;align-items:start;padding:.8rem .85rem;border:1px solid rgba(139,105,20,.28);background:rgba(255,255,255,.18)}
.spell input{margin-top:.18rem;accent-color:var(--rule)}
.map-quote{font-style:italic;font-size:1.15rem;line-height:1.7;padding:.2rem 0}
.footer{border-top:3px double var(--rule);margin-top:1rem;padding-top:.9rem;display:flex;justify-content:space-between;gap:1rem;flex-wrap:wrap;font-family:'Cinzel',serif;font-size:.75rem;letter-spacing:.08em;text-transform:uppercase;color:var(--ink3)}
@media (max-width:700px){.columns.two{grid-template-columns:1fr}.meta,.footer{flex-direction:column}.paper-inner{padding:1rem}.lead-quote{font-size:1.2rem}}
</style>
</head>
<body>
  <div id="motes" aria-hidden="true"></div>
  <main class="paper">
    <div class="paper-inner">
      <header class="nameplate">
        <div class="kicker">Wizarding Britain Morning Edition</div>
        <div class="flag">The Daily Prophet</div>
        <div class="sub">$(HtmlEncode $Data.MastheadLine)</div>
      </header>

      <div class="meta">
        <div>$(HtmlEncode $dateDisplay)</div>
        <div>Edition: $(HtmlEncode $Data.Edition)</div>
        <div>Issue ID: $(HtmlEncode $Data.IssueId)</div>
      </div>

      <section class="summary">
        <strong>Overall</strong>$(HtmlEncode $Data.Overall)
      </section>

      $(Render-FrontPage -FrontLead $Data.FrontLead -FrontColumns $Data.FrontColumns)

      $(Render-Watchlist -Items $Data.WatchlistItems)

      <section class="section">
        <div class="section-kicker">Potion Notes</div>
        <div class="note-box">$(HtmlEncode $Data.PotionNotes)</div>
      </section>

      $(Render-Spells -Todos $Data.Spells)

      <section class="section">
        <div class="section-kicker">The Map</div>
        <blockquote class="map-quote">$(HtmlEncode $Data.MapText)</blockquote>
      </section>

      <footer class="footer">
        <div>Sources: $(HtmlEncode $sourcesHtml)</div>
        <div>Filed under $(HtmlEncode $Data.Title)</div>
      </footer>
    </div>
  </main>
<script>
const wrap=document.getElementById('motes');
for(let i=0;i<24;i++){
  const mote=document.createElement('span');
  mote.className='mote';
  mote.style.left=Math.random()*100+'vw';
  mote.style.animationDuration=(10+Math.random()*14)+'s';
  mote.style.animationDelay=(-Math.random()*18)+'s';
  mote.style.setProperty('--dx',((Math.random()*50)-25)+'px');
  wrap.appendChild(mote);
}
</script>
</body>
</html>
"@
}

function Render-Placeholder([string]$IssueDate) {
  $dateDisplay = [datetime]::Parse($IssueDate).ToString('dddd, MMMM d, yyyy')
  return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>The Daily Prophet — $(HtmlEncode $dateDisplay) · Morning</title>
<link href="https://fonts.googleapis.com/css2?family=UnifrakturMaguntia&family=IM+Fell+English:ital@0;1&family=IM+Fell+English+SC&family=Cinzel:wght@400;600;900&family=Playfair+Display:ital,wght@0,400;0,700;1,400&display=swap" rel="stylesheet">
<style>
:root {
  --p:#f2e8d0;--p2:#ecdfc5;--p3:#d9c9a0;
  --ink:#1a1208;--ink2:#3d2f1a;--ink3:#6b5535;
  --gold:#b8902a;--gold3:#e8cc7a;--rule:#8b6914;
  --red:#8b1a1a;--red2:#5c0f0f;--sh:rgba(26,18,8,.32);
}
*{margin:0;padding:0;box-sizing:border-box;}
body{background:#18110a;min-height:100vh;padding:2rem 1rem 5rem;font-family:'IM Fell English',Georgia,serif;color:var(--ink);}
.paper{max-width:920px;margin:0 auto;background:var(--p);background-image:linear-gradient(158deg,#f6eed9 0%,#f2e8d0 40%,#ecdfc5 75%,#e4d5b5 100%);box-shadow:0 0 0 1px var(--p3),0 22px 60px var(--sh);padding:1.5rem 1.2rem 2rem}
.nameplate{border-top:3px double var(--rule);border-bottom:3px double var(--rule);padding:1rem 0 .85rem;text-align:center}
.kicker{font-family:'Cinzel',serif;font-size:.72rem;letter-spacing:.2em;text-transform:uppercase;color:var(--ink3);margin-bottom:.35rem}
.flag{font-family:'UnifrakturMaguntia',serif;font-size:clamp(2.8rem,8vw,5.2rem);line-height:1;color:var(--ink)}
.sub{font-family:'IM Fell English SC',serif;font-size:.9rem;letter-spacing:.08em;color:var(--ink2);margin-top:.45rem}
.notice{padding:2rem 0 1rem;text-align:center}
.notice h1{font-family:'Playfair Display',serif;font-size:2.2rem;margin-bottom:.65rem}
.notice p{font-size:1.15rem;line-height:1.6;color:var(--ink2)}
</style>
</head>
<body>
  <main class="paper">
    <header class="nameplate">
      <div class="kicker">Wizarding Britain Morning Edition</div>
      <div class="flag">The Daily Prophet</div>
      <div class="sub">No issue on the stands just yet</div>
    </header>
    <section class="notice">
      <h1>No issue yet</h1>
      <p>Check back after 6 AM ET for the next owl-posted morning edition.</p>
    </section>
  </main>
</body>
</html>
"@
}

$issueDate = Get-IssueDate
$queryBody = @{
  filter = @{
    and = @(
      @{ property = 'Date'; date = @{ equals = $issueDate } },
      @{ property = 'Edition'; select = @{ equals = 'Morning' } }
    )
  }
}

$resp = Invoke-NotionJson -Method Post -Uri "https://api.notion.com/v1/databases/$DatabaseId/query" -Body $queryBody

if (-not $resp.results -or $resp.results.Count -eq 0) {
  Render-Placeholder -IssueDate $issueDate | Set-Content -LiteralPath $OutputPath -Encoding UTF8
  Write-Host "Wrote placeholder issue for $issueDate"
  exit 0
}

$page = $resp.results[0]
$props = $page.properties
$blocks = @(Get-NotionChildren $page.id)

$masthead = ($blocks | Where-Object { $_.type -eq 'callout' } | Select-Object -First 1)
$mastheadLine = ''
if ($masthead -and $masthead.has_children) {
  $mastheadLine = ((Get-NotionChildren $masthead.id) | ForEach-Object { Get-BlockText $_ }) -join ' '
}

$frontLead = Get-BlockText ($blocks | Where-Object { $_.type -eq 'quote' } | Select-Object -First 1)
$columnLists = @($blocks | Where-Object { $_.type -eq 'column_list' })
$frontColumns = if ($columnLists.Count -gt 0) { Get-ColumnBlocks $columnLists[0].id } else { @() }
$watchColumns = if ($columnLists.Count -gt 1) { Get-ColumnBlocks $columnLists[1].id } else { @() }

$watchlistItems = @()
foreach ($column in $watchColumns) {
  foreach ($callout in $column) {
    if ($callout.type -eq 'callout') {
      $watchlistItems += Get-WatchlistItem $callout
    }
  }
}

$potionNotes = Get-BlockText ($blocks | Where-Object { $_.type -eq 'callout' } | Select-Object -Last 1)
$spells = @($blocks | Where-Object { $_.type -eq 'to_do' })
$mapText = Get-BlockText ($blocks | Where-Object { $_.type -eq 'quote' } | Select-Object -Last 1)

$data = @{
  Title        = Get-PropertyPlainText $props.Title
  Date         = Get-PropertyPlainText $props.Date
  Edition      = Get-PropertyPlainText $props.Edition
  IssueId      = Get-PropertyPlainText $props.'Issue ID'
  Overall      = Get-PropertyPlainText $props.Overall
  Sources      = Get-PropertyPlainText $props.Sources
  MastheadLine = $mastheadLine
  FrontLead    = $frontLead
  FrontColumns = $frontColumns
  WatchlistItems = $watchlistItems
  PotionNotes  = $potionNotes
  Spells       = $spells
  MapText      = $mapText
}

Render-Page -Data $data | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Host "Wrote issue to $OutputPath"
