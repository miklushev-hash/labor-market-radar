param(
  [string]$DigestPath,
  [ValidateSet('auto','short_update','director_deck')]
  [string]$DocumentKind = 'auto'
)

$ErrorActionPreference = 'Stop'

function Get-LatestDigestPath {
  param([string]$Root)
  $dataRoot = Join-Path $Root 'data'
  $candidate = Get-ChildItem -LiteralPath $dataRoot -Directory |
    Where-Object {
      $_.Name -match '^\d{4}-W\d{2}$' -or
      $_.Name -match '^\d{4}-\d{2}$' -or
      $_.Name -match '^\d{4}-Q[1-4]$' -or
      $_.Name -match '^\d{4}-H[12]$' -or
      $_.Name -match '^\d{4}$'
    } |
    Sort-Object Name -Descending |
    Select-Object -First 1
  if (-not $candidate) {
    throw 'No weekly, monthly, quarterly, halfyear, or annual period found in data.'
  }
  $weeklyDigest = Join-Path $candidate.FullName 'weekly_digest_multilayer.md'
  $monthlyDigest = Join-Path $candidate.FullName 'monthly_digest.md'
  $quarterlyDigest = Join-Path $candidate.FullName 'quarterly_digest.md'
  $halfyearDigest = Join-Path $candidate.FullName 'halfyear_digest.md'
  $annualDigest = Join-Path $candidate.FullName 'annual_digest.md'
  if (Test-Path -LiteralPath $weeklyDigest) {
    return $weeklyDigest
  }
  if (Test-Path -LiteralPath $monthlyDigest) {
    return $monthlyDigest
  }
  if (Test-Path -LiteralPath $quarterlyDigest) {
    return $quarterlyDigest
  }
  if (Test-Path -LiteralPath $halfyearDigest) {
    return $halfyearDigest
  }
  if (Test-Path -LiteralPath $annualDigest) {
    return $annualDigest
  }
  throw "No supported digest file found in $($candidate.FullName)."
}

function Escape-Html {
  param([string]$Text)
  if ($null -eq $Text) { return '' }
  return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Move-ToProjectTrash {
  param(
    [string]$Path,
    [string]$Root
  )
  if (-not (Test-Path -LiteralPath $Path)) { return }

  # Windows PowerShell 5 may decode UTF-8 scripts without BOM as ANSI.
  $trashName = -join ([char[]](0x041A,0x043E,0x0440,0x0437,0x0438,0x043D,0x0430))
  $trashRoot = Join-Path $Root $trashName
  $trashSubdir = Join-Path $trashRoot 'replaced_presentations'
  if (-not (Test-Path -LiteralPath $trashSubdir)) {
    New-Item -ItemType Directory -Path $trashSubdir -Force | Out-Null
  }

  $item = Get-Item -LiteralPath $Path
  $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $targetName = '{0}__{1}{2}' -f $item.BaseName, $stamp, $item.Extension
  Move-Item -LiteralPath $item.FullName -Destination (Join-Path $trashSubdir $targetName) -Force
}

function Convert-Inline {
  param([string]$Text)
  $encoded = Escape-Html $Text
  return [System.Text.RegularExpressions.Regex]::Replace(
    $encoded,
    '`([^`]+)`',
    '<code>$1</code>'
  )
}

function Convert-TableBlock {
  param([string[]]$Lines)
  if ($Lines.Count -lt 2) { return '' }

  $rows = @()
  foreach ($line in $Lines) {
    $trimmed = $line.Trim()
    if (-not $trimmed.StartsWith('|')) { continue }
    $cells = $trimmed.Trim('|').Split('|') | ForEach-Object { $_.Trim() }
    $rows += ,$cells
  }

  if ($rows.Count -lt 2) { return '' }

  $header = $rows[0]
  $bodyRows = @()
  for ($i = 2; $i -lt $rows.Count; $i++) {
    $bodyRows += ,$rows[$i]
  }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine('<div class="table-wrap"><table><thead><tr>')
  foreach ($cell in $header) {
    [void]$sb.AppendLine("<th>$(Convert-Inline $cell)</th>")
  }
  [void]$sb.AppendLine('</tr></thead><tbody>')
  foreach ($row in $bodyRows) {
    [void]$sb.AppendLine('<tr>')
    foreach ($cell in $row) {
      [void]$sb.AppendLine("<td>$(Convert-Inline $cell)</td>")
    }
    [void]$sb.AppendLine('</tr>')
  }
  [void]$sb.AppendLine('</tbody></table></div>')
  return $sb.ToString()
}

function Convert-MarkdownToHtmlBody {
  param([string]$Markdown)

  $lines = $Markdown -split "`r?`n"
  $sb = New-Object System.Text.StringBuilder
  $inUl = $false
  $inOl = $false
  $paragraph = New-Object System.Collections.Generic.List[string]

  function Flush-Paragraph {
    param($Paragraph, $Builder)
    if ($Paragraph.Count -gt 0) {
      $text = ($Paragraph -join ' ').Trim()
      if ($text.Length -gt 0) {
        [void]$Builder.AppendLine("<p>$(Convert-Inline $text)</p>")
      }
      $Paragraph.Clear()
    }
  }

  function Close-Lists {
    param([ref]$Ul, [ref]$Ol, $Builder)
    if ($Ul.Value) {
      [void]$Builder.AppendLine('</ul>')
      $Ul.Value = $false
    }
    if ($Ol.Value) {
      [void]$Builder.AppendLine('</ol>')
      $Ol.Value = $false
    }
  }

  $i = 0
  while ($i -lt $lines.Count) {
    $line = $lines[$i]
    $trimmed = $line.Trim()

    if ($trimmed -eq '') {
      Flush-Paragraph -Paragraph $paragraph -Builder $sb
      Close-Lists ([ref]$inUl) ([ref]$inOl) $sb
      $i++
      continue
    }

    if ($trimmed.StartsWith('|')) {
      Flush-Paragraph -Paragraph $paragraph -Builder $sb
      Close-Lists ([ref]$inUl) ([ref]$inOl) $sb
      $tableLines = New-Object System.Collections.Generic.List[string]
      while ($i -lt $lines.Count -and $lines[$i].Trim().StartsWith('|')) {
        $tableLines.Add($lines[$i])
        $i++
      }
      [void]$sb.AppendLine((Convert-TableBlock $tableLines.ToArray()))
      continue
    }

    if ($trimmed -match '^###\s+(.+)$') {
      Flush-Paragraph -Paragraph $paragraph -Builder $sb
      Close-Lists ([ref]$inUl) ([ref]$inOl) $sb
      [void]$sb.AppendLine("<h3>$(Convert-Inline $matches[1])</h3>")
      $i++
      continue
    }

    if ($trimmed -match '^##\s+(.+)$') {
      Flush-Paragraph -Paragraph $paragraph -Builder $sb
      Close-Lists ([ref]$inUl) ([ref]$inOl) $sb
      [void]$sb.AppendLine("<h2>$(Convert-Inline $matches[1])</h2>")
      $i++
      continue
    }

    if ($trimmed -match '^#\s+(.+)$') {
      Flush-Paragraph -Paragraph $paragraph -Builder $sb
      Close-Lists ([ref]$inUl) ([ref]$inOl) $sb
      [void]$sb.AppendLine("<h1>$(Convert-Inline $matches[1])</h1>")
      $i++
      continue
    }

    if ($trimmed -match '^- (.+)$') {
      Flush-Paragraph -Paragraph $paragraph -Builder $sb
      if ($inOl) {
        [void]$sb.AppendLine('</ol>')
        $inOl = $false
      }
      if (-not $inUl) {
        [void]$sb.AppendLine('<ul>')
        $inUl = $true
      }
      [void]$sb.AppendLine("<li>$(Convert-Inline $matches[1])</li>")
      $i++
      continue
    }

    if ($trimmed -match '^\d+\.\s+(.+)$') {
      Flush-Paragraph -Paragraph $paragraph -Builder $sb
      if ($inUl) {
        [void]$sb.AppendLine('</ul>')
        $inUl = $false
      }
      if (-not $inOl) {
        [void]$sb.AppendLine('<ol>')
        $inOl = $true
      }
      [void]$sb.AppendLine("<li>$(Convert-Inline $matches[1])</li>")
      $i++
      continue
    }

    $paragraph.Add($trimmed)
    $i++
  }

  Flush-Paragraph -Paragraph $paragraph -Builder $sb
  Close-Lists ([ref]$inUl) ([ref]$inOl) $sb
  return $sb.ToString()
}

function Build-HtmlDocument {
  param(
    [string]$Title,
    [string]$Subtitle,
    [string]$PeriodLabel,
    [string]$DocumentKind,
    [string]$SourceName,
    [string]$BodyHtml
  )

  $kindLabel = if ($DocumentKind -eq 'short_update') { '&#1050;&#1086;&#1088;&#1086;&#1090;&#1082;&#1080;&#1081; &#1074;&#1099;&#1087;&#1091;&#1089;&#1082;' } else { '&#1044;&#1080;&#1088;&#1077;&#1082;&#1090;&#1086;&#1088;&#1089;&#1082;&#1072;&#1103; &#1074;&#1077;&#1088;&#1089;&#1080;&#1103;' }

@"
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>$Title</title>
  <style>
    :root {
      --ink: #16202a;
      --muted: #5f6b76;
      --line: #d8dee6;
      --paper: #f3f5f8;
      --white: #ffffff;
      --blue: #184e9e;
      --shadow: 0 16px 36px rgba(15, 23, 42, 0.10);
    }
    * { box-sizing: border-box; }
    html, body { margin: 0; padding: 0; }
    body { font-family: "Segoe UI", Arial, sans-serif; color: var(--ink); background: var(--paper); line-height: 1.55; }
    .page { width: min(980px, calc(100vw - 32px)); margin: 28px auto 40px; background: var(--white); border: 1px solid var(--line); border-radius: 10px; box-shadow: var(--shadow); overflow: hidden; }
    .topbar { display: flex; justify-content: space-between; align-items: center; gap: 16px; padding: 14px 28px; background: #eef3f9; border-bottom: 1px solid var(--line); font-size: 14px; color: var(--muted); }
    .topbar button { border: 1px solid var(--line); background: var(--white); color: var(--ink); border-radius: 6px; padding: 8px 12px; font: inherit; cursor: pointer; }
    .hero { padding: 42px 42px 24px; background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%); }
    .eyebrow { margin: 0 0 10px; color: var(--blue); font-size: 13px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; }
    h1 { margin: 0 0 12px; font-size: 38px; line-height: 1.08; }
    h2 { margin: 30px 0 14px; font-size: 27px; line-height: 1.18; }
    h3 { margin: 20px 0 10px; font-size: 20px; line-height: 1.25; }
    .subtitle { max-width: 780px; margin: 0; color: var(--muted); font-size: 19px; }
    .meta { display: flex; gap: 12px; flex-wrap: wrap; padding: 0 42px 20px; }
    .meta span { display: inline-flex; align-items: center; padding: 6px 10px; border: 1px solid var(--line); border-radius: 999px; background: #f8fafc; color: var(--muted); font-size: 13px; }
    .content { padding: 0 42px 34px; }
    p { margin: 0 0 12px; font-size: 17px; }
    ul, ol { margin: 0 0 14px; padding-left: 24px; }
    li { margin: 0 0 9px; font-size: 17px; }
    code { background: #eef3f9; border-radius: 4px; padding: 1px 6px; font-family: Consolas, monospace; font-size: 0.95em; }
    .table-wrap { border: 1px solid var(--line); border-radius: 8px; overflow: hidden; margin: 14px 0 18px; }
    table { width: 100%; border-collapse: collapse; font-size: 16px; }
    th, td { text-align: left; vertical-align: top; padding: 12px 14px; border-bottom: 1px solid var(--line); }
    th { background: #eef3f9; font-size: 14px; text-transform: uppercase; letter-spacing: 0.04em; color: #334155; }
    tr:last-child td { border-bottom: 0; }
    .footer { padding: 18px 42px 28px; border-top: 1px solid var(--line); color: var(--muted); font-size: 14px; }
    @media (max-width: 840px) {
      .page { width: min(100vw - 16px, 980px); margin: 10px auto 18px; }
      .hero, .content, .footer { padding-left: 20px; padding-right: 20px; }
      .meta { padding-left: 20px; padding-right: 20px; }
      h1 { font-size: 30px; }
      h2 { font-size: 24px; }
      .topbar { padding: 12px 20px; }
    }
    @media print {
      body { background: #ffffff; }
      .page { width: 100%; margin: 0; border: 0; border-radius: 0; box-shadow: none; }
      .topbar { display: none; }
    }
  </style>
</head>
<body>
  <main class="page">
    <div class="topbar">
      <strong>$kindLabel | $PeriodLabel</strong>
      <button type="button" onclick="window.print()">&#1055;&#1077;&#1095;&#1072;&#1090;&#1100; / PDF</button>
    </div>
    <section class="hero">
      <p class="eyebrow">$kindLabel</p>
      <h1>$(Escape-Html $Title)</h1>
      <p class="subtitle">$(Escape-Html $Subtitle)</p>
    </section>
    <section class="meta">
      <span>&#1055;&#1077;&#1088;&#1080;&#1086;&#1076;: $PeriodLabel</span>
      <span>&#1048;&#1089;&#1090;&#1086;&#1095;&#1085;&#1080;&#1082;: $SourceName</span>
      <span>&#1060;&#1086;&#1088;&#1084;&#1072;&#1090;: HTML + PDF</span>
    </section>
    <section class="content">
      $BodyHtml
    </section>
    <section class="footer">
      &#1044;&#1086;&#1082;&#1091;&#1084;&#1077;&#1085;&#1090; &#1089;&#1086;&#1073;&#1088;&#1072;&#1085; &#1072;&#1074;&#1090;&#1086;&#1084;&#1072;&#1090;&#1080;&#1095;&#1077;&#1089;&#1082;&#1080; &#1080;&#1079; Markdown-&#1092;&#1072;&#1081;&#1083;&#1072; &#1086;&#1073;&#1079;&#1086;&#1088;&#1072;.
    </section>
  </main>
</body>
</html>
"@
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = if (Test-Path -LiteralPath (Join-Path $scriptDir 'data')) {
  $scriptDir
} elseif (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $scriptDir) 'data')) {
  Split-Path -Parent $scriptDir
} else {
  $scriptDir
}
if (-not $DigestPath) {
  $DigestPath = Get-LatestDigestPath -Root $root
}

$resolvedDigest = (Resolve-Path -LiteralPath $DigestPath).Path
$markdown = Get-Content -LiteralPath $resolvedDigest -Raw -Encoding UTF8

$sourceName = Split-Path $resolvedDigest -Leaf
$periodLabelMatch = [regex]::Match($markdown, 'period_label`:\s*`([^`]+)`')
$weekLabelMatch = [regex]::Match($markdown, 'week_label`:\s*`([^`]+)`')
$periodLabel = if ($periodLabelMatch.Success) {
  $periodLabelMatch.Groups[1].Value
} elseif ($weekLabelMatch.Success) {
  $weekLabelMatch.Groups[1].Value
} else {
  Split-Path (Split-Path $resolvedDigest -Parent) -Leaf
}
$isMonthly = $sourceName -eq 'monthly_digest.md'
$isQuarterly = $sourceName -eq 'quarterly_digest.md'
$isHalfyear = $sourceName -eq 'halfyear_digest.md'
$isAnnual = $sourceName -eq 'annual_digest.md'

$titleMatch = [regex]::Match($markdown, '(?m)^#\s+(.+)$')
$title = if ($titleMatch.Success) {
  $titleMatch.Groups[1].Value.Trim()
} elseif ($isAnnual) {
  "&#1043;&#1086;&#1076;&#1086;&#1074;&#1086;&#1081; &#1074;&#1099;&#1087;&#1091;&#1089;&#1082; $periodLabel"
} elseif ($isHalfyear) {
  "&#1055;&#1086;&#1083;&#1091;&#1075;&#1086;&#1076;&#1086;&#1074;&#1086;&#1081; &#1074;&#1099;&#1087;&#1091;&#1089;&#1082; $periodLabel"
} elseif ($isQuarterly) {
  "&#1050;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;&#1100;&#1085;&#1099;&#1081; &#1074;&#1099;&#1087;&#1091;&#1089;&#1082; $periodLabel"
} elseif ($isMonthly) {
  "&#1052;&#1077;&#1089;&#1103;&#1095;&#1085;&#1099;&#1081; &#1074;&#1099;&#1087;&#1091;&#1089;&#1082; $periodLabel"
} else {
  "&#1053;&#1077;&#1076;&#1077;&#1083;&#1100;&#1085;&#1099;&#1081; &#1074;&#1099;&#1087;&#1091;&#1089;&#1082; $periodLabel"
}

$h2Count = ([regex]::Matches($markdown, '(?m)^##\s+')).Count
$resolvedKind = switch ($DocumentKind) {
  'short_update' { 'short_update' }
  'director_deck' { 'director_deck' }
  default {
    if ($h2Count -le 8) { 'short_update' } else { 'director_deck' }
  }
}

$subtitle = if ($resolvedKind -eq 'short_update') {
  if ($isAnnual) {
    '&#1050;&#1086;&#1088;&#1086;&#1090;&#1082;&#1080;&#1081; &#1075;&#1086;&#1076;&#1086;&#1074;&#1086;&#1081; &#1074;&#1099;&#1087;&#1091;&#1089;&#1082; &#1087;&#1086; &#1087;&#1086;&#1076;&#1090;&#1074;&#1077;&#1088;&#1078;&#1076;&#1077;&#1085;&#1085;&#1086;&#1084;&#1091; &#1082;&#1086;&#1088;&#1087;&#1091;&#1089;&#1091; &#1075;&#1086;&#1076;&#1072;.'
  } elseif ($isHalfyear) {
    '&#1050;&#1086;&#1088;&#1086;&#1090;&#1082;&#1080;&#1081; &#1087;&#1086;&#1083;&#1091;&#1075;&#1086;&#1076;&#1086;&#1074;&#1086;&#1081; &#1074;&#1099;&#1087;&#1091;&#1089;&#1082; &#1087;&#1086; &#1087;&#1086;&#1076;&#1090;&#1074;&#1077;&#1088;&#1078;&#1076;&#1077;&#1085;&#1085;&#1086;&#1084;&#1091; &#1082;&#1086;&#1088;&#1087;&#1091;&#1089;&#1091; &#1087;&#1077;&#1088;&#1080;&#1086;&#1076;&#1072;.'
  } elseif ($isQuarterly) {
    '&#1050;&#1086;&#1088;&#1086;&#1090;&#1082;&#1080;&#1081; &#1082;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;&#1100;&#1085;&#1099;&#1081; &#1074;&#1099;&#1087;&#1091;&#1089;&#1082; &#1087;&#1086; &#1087;&#1086;&#1076;&#1090;&#1074;&#1077;&#1088;&#1078;&#1076;&#1077;&#1085;&#1085;&#1086;&#1084;&#1091; &#1082;&#1086;&#1088;&#1087;&#1091;&#1089;&#1091; &#1087;&#1077;&#1088;&#1080;&#1086;&#1076;&#1072;.'
  } elseif ($isMonthly) {
    '&#1050;&#1086;&#1088;&#1086;&#1090;&#1082;&#1080;&#1081; &#1084;&#1077;&#1089;&#1103;&#1095;&#1085;&#1099;&#1081; &#1074;&#1099;&#1087;&#1091;&#1089;&#1082; &#1087;&#1086; &#1087;&#1086;&#1076;&#1090;&#1074;&#1077;&#1088;&#1078;&#1076;&#1077;&#1085;&#1085;&#1086;&#1084;&#1091; &#1082;&#1086;&#1088;&#1087;&#1091;&#1089;&#1091; &#1084;&#1077;&#1089;&#1103;&#1094;&#1072;.'
  } else {
    '&#1050;&#1086;&#1088;&#1086;&#1090;&#1082;&#1080;&#1081; &#1085;&#1077;&#1076;&#1077;&#1083;&#1100;&#1085;&#1099;&#1081; &#1074;&#1099;&#1087;&#1091;&#1089;&#1082; &#1087;&#1086; &#1087;&#1086;&#1076;&#1090;&#1074;&#1077;&#1088;&#1078;&#1076;&#1077;&#1085;&#1085;&#1086;&#1084;&#1091; &#1082;&#1086;&#1088;&#1087;&#1091;&#1089;&#1091; &#1085;&#1077;&#1076;&#1077;&#1083;&#1080;.'
  }
} else {
  if ($isAnnual) {
    '&#1055;&#1086;&#1083;&#1085;&#1072;&#1103; &#1075;&#1086;&#1076;&#1086;&#1074;&#1072;&#1103; &#1074;&#1077;&#1088;&#1089;&#1080;&#1103; &#1076;&#1083;&#1103; &#1095;&#1090;&#1077;&#1085;&#1080;&#1103;, &#1087;&#1077;&#1095;&#1072;&#1090;&#1080; &#1080; &#1086;&#1073;&#1089;&#1091;&#1078;&#1076;&#1077;&#1085;&#1080;&#1103;.'
  } elseif ($isHalfyear) {
    '&#1055;&#1086;&#1083;&#1085;&#1072;&#1103; &#1087;&#1086;&#1083;&#1091;&#1075;&#1086;&#1076;&#1086;&#1074;&#1072;&#1103; &#1074;&#1077;&#1088;&#1089;&#1080;&#1103; &#1076;&#1083;&#1103; &#1095;&#1090;&#1077;&#1085;&#1080;&#1103;, &#1087;&#1077;&#1095;&#1072;&#1090;&#1080; &#1080; &#1086;&#1073;&#1089;&#1091;&#1078;&#1076;&#1077;&#1085;&#1080;&#1103;.'
  } elseif ($isQuarterly) {
    '&#1055;&#1086;&#1083;&#1085;&#1072;&#1103; &#1082;&#1074;&#1072;&#1088;&#1090;&#1072;&#1083;&#1100;&#1085;&#1072;&#1103; &#1074;&#1077;&#1088;&#1089;&#1080;&#1103; &#1076;&#1083;&#1103; &#1095;&#1090;&#1077;&#1085;&#1080;&#1103;, &#1087;&#1077;&#1095;&#1072;&#1090;&#1080; &#1080; &#1086;&#1073;&#1089;&#1091;&#1078;&#1076;&#1077;&#1085;&#1080;&#1103;.'
  } elseif ($isMonthly) {
    '&#1055;&#1086;&#1083;&#1085;&#1072;&#1103; &#1084;&#1077;&#1089;&#1103;&#1095;&#1085;&#1072;&#1103; &#1074;&#1077;&#1088;&#1089;&#1080;&#1103; &#1076;&#1083;&#1103; &#1095;&#1090;&#1077;&#1085;&#1080;&#1103;, &#1087;&#1077;&#1095;&#1072;&#1090;&#1080; &#1080; &#1086;&#1073;&#1089;&#1091;&#1078;&#1076;&#1077;&#1085;&#1080;&#1103;.'
  } else {
    '&#1055;&#1086;&#1083;&#1085;&#1072;&#1103; &#1085;&#1077;&#1076;&#1077;&#1083;&#1100;&#1085;&#1072;&#1103; &#1074;&#1077;&#1088;&#1089;&#1080;&#1103; &#1076;&#1083;&#1103; &#1095;&#1090;&#1077;&#1085;&#1080;&#1103;, &#1087;&#1077;&#1095;&#1072;&#1090;&#1080; &#1080; &#1086;&#1073;&#1089;&#1091;&#1078;&#1076;&#1077;&#1085;&#1080;&#1103;.'
  }
}

$presentations = Join-Path $root 'presentations'
if (-not (Test-Path -LiteralPath $presentations)) {
  New-Item -ItemType Directory -Path $presentations | Out-Null
}

$baseName = if ($resolvedKind -eq 'short_update') {
  if ($isAnnual) {
    "${periodLabel}_annual_brief"
  } elseif ($isHalfyear) {
    "${periodLabel}_halfyear_brief"
  } elseif ($isQuarterly) {
    "${periodLabel}_quarterly_brief"
  } elseif ($isMonthly) {
    "${periodLabel}_monthly_brief"
  } else {
    "${periodLabel}_short_update_brief"
  }
} else {
  if ($isAnnual) {
    "${periodLabel}_annual_director_deck"
  } elseif ($isHalfyear) {
    "${periodLabel}_halfyear_director_deck"
  } elseif ($isQuarterly) {
    "${periodLabel}_quarterly_director_deck"
  } elseif ($isMonthly) {
    "${periodLabel}_monthly_director_deck"
  } else {
    "${periodLabel}_director_deck"
  }
}

$htmlPath = Join-Path $presentations ($baseName + '.html')
$pdfPath = Join-Path $presentations ($baseName + '.pdf')

if (Test-Path -LiteralPath $htmlPath) {
  Move-ToProjectTrash -Path $htmlPath -Root $root
}
if (Test-Path -LiteralPath $pdfPath) {
  Move-ToProjectTrash -Path $pdfPath -Root $root
}

$bodyHtml = Convert-MarkdownToHtmlBody -Markdown $markdown
$fullHtml = Build-HtmlDocument -Title $title -Subtitle $subtitle -PeriodLabel $periodLabel -DocumentKind $resolvedKind -SourceName $sourceName -BodyHtml $bodyHtml
Set-Content -LiteralPath $htmlPath -Value $fullHtml -Encoding UTF8

$edgeCandidates = @(
  'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
  'C:\Program Files\Microsoft\Edge\Application\msedge.exe'
)
$edgePath = $edgeCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $edgePath) {
  throw 'Microsoft Edge not found.'
}

$uri = [System.Uri]$htmlPath
& $edgePath --headless --disable-gpu "--print-to-pdf=$pdfPath" $uri.AbsoluteUri | Out-Null

if (-not (Test-Path -LiteralPath $pdfPath)) {
  throw 'PDF was not created.'
}

[pscustomobject]@{
  PeriodLabel = $periodLabel
  DocumentKind = $resolvedKind
  DigestPath = $resolvedDigest
  HtmlPath = $htmlPath
  PdfPath = $pdfPath
} | ConvertTo-Json -Compress
