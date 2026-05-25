Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ErrorActionPreference = 'Stop'
$Root = Get-Location
$StoragePath = Join-Path $Root 'Хранилище.xlsx'
$SourcesPath = Join-Path $Root 'Источники данных.xlsx'
$Today = '2026-05-14'

$HeadersBase = @(
  'Период',
  'Название источника',
  'дата сбора информации',
  'тип источника: сайт / telegram / иное',
  'заголовок конкретного материала / статус отсутствия данных',
  'что собираем',
  'тема предмета',
  'регион',
  'отрасль',
  'тип сигнала',
  'тональность',
  'важность',
  'контур',
  'использован в обзоре',
  'причина включения',
  'статус проверки источника',
  'уровень доказательности',
  'ссылка',
  'комментарий',
  'стратегический сигнал',
  'объект стратегического сигнала',
  'влияет на рынок труда',
  'влияет на ИТ-рынок',
  'влияет на значимых работодателей',
  'влияет на стратегических клиентов',
  'сдвиг в логике отрасли',
  'механизм влияния',
  'шум или общий сдвиг'
)

function Get-ZipEntryByName($Zip, [string]$Name) {
  $Alt = $Name -replace '/', '\'
  return $Zip.Entries | Where-Object { $_.FullName -eq $Name -or $_.FullName -eq $Alt } | Select-Object -First 1
}

function Read-ZipText($Zip, [string]$Name) {
  $Entry = Get-ZipEntryByName $Zip $Name
  if (-not $Entry) { throw "Missing zip entry: $Name" }
  $Reader = [IO.StreamReader]::new($Entry.Open(), [Text.Encoding]::UTF8)
  try { return $Reader.ReadToEnd() } finally { $Reader.Close() }
}

function Get-CellText($Cell, $SharedStrings) {
  $Value = [string]$Cell.v
  if ($null -eq $Value -or $Value -eq '') { return '' }
  if ($Cell.t -eq 's') { return $SharedStrings[[int]$Value] }
  return $Value
}

function Get-ColNum([string]$Ref) {
  $Letters = $Ref -replace '\d',''
  $N = 0
  foreach ($Ch in $Letters.ToCharArray()) {
    $N = $N * 26 + ([int][char]$Ch - [int][char]'A' + 1)
  }
  return $N
}

function Get-ColLetters([int]$Num) {
  $S = ''
  while ($Num -gt 0) {
    $Num--
    $S = [char](65 + ($Num % 26)) + $S
    $Num = [math]::Floor($Num / 26)
  }
  return $S
}

function Normalize-SourceName([string]$Name) {
  if (-not $Name) { return '' }
  $N = $Name -replace '\[[^\]]+\]\([^)]+\)', '$0'
  $N = $N -replace '^\s*\[([^\]]+)\]\([^)]+\)\s*$', '$1'
  $N = $N -replace '<[^>]+>', ''
  $N = $N -replace '\s+', ' '
  $N = $N.Trim()
  $N = $N -replace '\s*[:：].*$', ''
  return $N.Trim().ToLowerInvariant()
}

function Split-MarkdownRow([string]$Line) {
  $T = $Line.Trim()
  if ($T.StartsWith('|')) { $T = $T.Substring(1) }
  if ($T.EndsWith('|')) { $T = $T.Substring(0, $T.Length - 1) }
  return @($T -split '\|' | ForEach-Object { $_.Trim() })
}

function Get-MarkdownLinkInfo([string]$Cell) {
  $Name = ($Cell -replace '<[^>]+>', '').Trim()
  $Url = ''
  if ($Cell -match '\[([^\]]+)\]\(([^)]+)\)') {
    $Name = $Matches[1].Trim()
    $Url = $Matches[2].Trim()
  }
  return [PSCustomObject]@{ Name = $Name; Url = $Url }
}

function Parse-ReviewTables([string]$Path) {
  $Rows = @()
  if (-not (Test-Path -LiteralPath $Path)) { return $Rows }
  $Lines = Get-Content -LiteralPath $Path -Encoding UTF8
  for ($I = 0; $I -lt $Lines.Count - 1; $I++) {
    if ($Lines[$I].Trim().StartsWith('|') -and $Lines[$I + 1] -match '^\s*\|?\s*:?-{3,}') {
      $Headers = Split-MarkdownRow $Lines[$I]
      $SourceIdx = [Array]::FindIndex($Headers, [Predicate[string]]{ param($H) $H -match 'Источник' })
      if ($SourceIdx -ne 0) { continue }
      $StatusIdx = [Array]::FindIndex($Headers, [Predicate[string]]{ param($H) $H -match 'Статус' })
      $WhatIdx = [Array]::FindIndex($Headers, [Predicate[string]]{ param($H) $H -match 'Что' })
      $UsedIdx = [Array]::FindIndex($Headers, [Predicate[string]]{ param($H) $H -match 'Использован' })
      $J = $I + 2
      while ($J -lt $Lines.Count -and $Lines[$J].Trim().StartsWith('|')) {
        $Cells = Split-MarkdownRow $Lines[$J]
        if ($Cells.Count -gt $SourceIdx) {
          $Link = Get-MarkdownLinkInfo $Cells[$SourceIdx]
          if ($Link.Name -and $Link.Name -notmatch '^-+$') {
            $Status = if ($StatusIdx -ge 0 -and $Cells.Count -gt $StatusIdx) { $Cells[$StatusIdx] } else { 'использовано в выпуске' }
            $What = if ($WhatIdx -ge 0 -and $Cells.Count -gt $WhatIdx) { $Cells[$WhatIdx] } else { '' }
            $Used = if ($UsedIdx -ge 0 -and $Cells.Count -gt $UsedIdx) { $Cells[$UsedIdx] } else { 'да' }
            if ($Status -match 'не использовано' -or $Used -match 'нет') { $Used = 'нет' }
            elseif ($Used -notmatch 'нет') { $Used = 'да' }
            $Rows += [PSCustomObject]@{
              Source = $Link.Name
              Key = Normalize-SourceName $Link.Name
              Status = $Status
              What = $What
              Used = $Used
              Url = $Link.Url
            }
          }
        }
        $J++
      }
    }
  }
  return $Rows
}

function Read-SourceRegistry {
  $Zip = [System.IO.Compression.ZipFile]::OpenRead($SourcesPath)
  try {
    [xml]$Wb = Read-ZipText $Zip 'xl/workbook.xml'
    [xml]$Rels = Read-ZipText $Zip 'xl/_rels/workbook.xml.rels'
    [xml]$Sst = Read-ZipText $Zip 'xl/sharedStrings.xml'
    $Shared = @()
    foreach ($Si in $Sst.sst.si) {
      if ($Si.t) { $Shared += [string]$Si.t }
      else { $Shared += (($Si.r | ForEach-Object { [string]$_.t }) -join '') }
    }
    $RelMap = @{}
    foreach ($R in $Rels.Relationships.Relationship) { $RelMap[$R.Id] = $R.Target }
    $SkipSheets = @('Доступность_источников', 'Частично_доступные', 'детали проблем')
    $All = @()
    foreach ($Sheet in $Wb.workbook.sheets.sheet) {
      $SheetName = [string]$Sheet.name
      if ($SkipSheets -contains $SheetName) { continue }
      $Rid = $Sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
      [xml]$Sh = Read-ZipText $Zip ('xl/' + $RelMap[$Rid])
      $Rows = @($Sh.worksheet.sheetData.row)
      if ($Rows.Count -eq 0) { continue }
      $HeaderByName = @{}
      foreach ($C in $Rows[0].c) {
        $HeaderByName[(Get-CellText $C $Shared)] = Get-ColNum $C.r
      }
      $NameCol = if ($HeaderByName.ContainsKey('Название')) { $HeaderByName['Название'] } elseif ($HeaderByName.ContainsKey('Название источника')) { $HeaderByName['Название источника'] } else { 1 }
      $CollectCol = if ($HeaderByName.ContainsKey('Что собираем')) { $HeaderByName['Что собираем'] } else { $NameCol + 1 }
      $UrlCol = if ($HeaderByName.ContainsKey('ссылка')) { $HeaderByName['ссылка'] } elseif ($HeaderByName.ContainsKey('Ссылка')) { $HeaderByName['Ссылка'] } else { $NameCol + 2 }
      foreach ($R in $Rows | Select-Object -Skip 1) {
        $Vals = @{}
        foreach ($C in $R.c) { $Vals[(Get-ColNum $C.r)] = Get-CellText $C $Shared }
        $Name = [string]$Vals[$NameCol]
        if (-not $Name -or $Name -eq 'Название') { continue }
        $All += [PSCustomObject]@{
          Source = $Name.Trim()
          Key = Normalize-SourceName $Name
          Collect = ([string]$Vals[$CollectCol]).Trim()
          Url = ([string]$Vals[$UrlCol]).Trim()
          Type = $SheetName
        }
      }
    }
    return $All | Sort-Object Source -Unique
  } finally {
    $Zip.Dispose()
  }
}

function Merge-ReviewRows($Rows) {
  $Map = @{}
  foreach ($R in $Rows) {
    $Key = $R.Key
    if (-not $Key -or $Key -eq 'остальные источники пилотного корпуса') { continue }
    if (-not $Map.ContainsKey($Key)) {
      $Map[$Key] = [PSCustomObject]@{
        Source = $R.Source
        Statuses = New-Object System.Collections.Generic.List[string]
        Whats = New-Object System.Collections.Generic.List[string]
        UsedValues = New-Object System.Collections.Generic.List[string]
        Urls = New-Object System.Collections.Generic.List[string]
      }
    }
    if ($R.Status) { $Map[$Key].Statuses.Add($R.Status) }
    if ($R.What) { $Map[$Key].Whats.Add($R.What) }
    if ($R.Used) { $Map[$Key].UsedValues.Add($R.Used) }
    if ($R.Url) { $Map[$Key].Urls.Add($R.Url) }
  }
  return $Map
}

function Build-StorageRows($Period, $Kind, $ReviewPath, $PresentationBase, $Registry) {
  $Parsed = Merge-ReviewRows (Parse-ReviewTables $ReviewPath)
  $Rows = [System.Collections.ArrayList]::new()
  $Seen = @{}
  foreach ($Src in $Registry) {
    $ParsedRow = $null
    if ($Parsed.ContainsKey($Src.Key)) { $ParsedRow = $Parsed[$Src.Key] }
    $Status = 'не восстановлено из готового обзора; требуется ручная проверка источника'
    $What = 'статус по этому источнику отсутствует в готовом обзоре'
    $Used = 'нет'
    $Reason = ''
    $Evidence = 'не восстановлено'
    $Link = $Src.Url
    if ($ParsedRow) {
      $Status = (($ParsedRow.Statuses | Select-Object -Unique) -join '; ')
      $What = (($ParsedRow.Whats | Select-Object -Unique) -join '; ')
      if (-not $What) { $What = 'источник указан в приложении готового обзора' }
      $Used = if (($ParsedRow.UsedValues | Where-Object { $_ -eq 'да' }).Count -gt 0) { 'да' } else { 'нет' }
      $Reason = if ($Used -eq 'да') { $What } else { '' }
      $Evidence = 'восстановлено из готового обзора'
      if (($ParsedRow.Urls | Where-Object { $_ } | Select-Object -First 1)) {
        $Link = (($ParsedRow.Urls | Where-Object { $_ } | Select-Object -Unique) -join '; ')
      }
      $Seen[$Src.Key] = $true
    }
    $Comment = "Восстановлено 2026-05-14 из $ReviewPath"
    if ($PresentationBase) { $Comment += "; презентации: $PresentationBase.html / $PresentationBase.pdf" }
    [void]$Rows.Add(
      [object[]]@(
        $Period,
        $Src.Source,
        $Today,
        $Src.Type,
        $What,
        $Src.Collect,
        'рынок труда / ИТ / экономика',
        'Россия',
        'рынок труда; ИТ',
        'не восстановлено из готового обзора',
        'не восстановлено из готового обзора',
        'не восстановлено из готового обзора',
        $Src.Type,
        $Used,
        $Reason,
        $Status,
        $Evidence,
        $Link,
        $Comment,
        'не восстановлено из готового обзора',
        '',
        'не восстановлено из готового обзора',
        'не восстановлено из готового обзора',
        'не восстановлено из готового обзора',
        'не восстановлено из готового обзора',
        'не восстановлено из готового обзора',
        'не восстановлено из готового обзора',
        'не восстановлено из готового обзора'
      )
    )
  }
  foreach ($Key in $Parsed.Keys) {
    if ($Seen.ContainsKey($Key)) { continue }
    $P = $Parsed[$Key]
    $What = (($P.Whats | Select-Object -Unique) -join '; ')
    $Status = (($P.Statuses | Select-Object -Unique) -join '; ')
    $Used = if (($P.UsedValues | Where-Object { $_ -eq 'да' }).Count -gt 0) { 'да' } else { 'нет' }
    $Link = (($P.Urls | Where-Object { $_ } | Select-Object -Unique) -join '; ')
    [void]$Rows.Add([object[]]@(
      $Period, $P.Source, $Today, 'источник вне текущего реестра', $What,
      '', 'рынок труда / ИТ / экономика', 'Россия', 'рынок труда; ИТ',
      'не восстановлено из готового обзора', 'не восстановлено из готового обзора',
      'не восстановлено из готового обзора', 'вне текущего реестра', $Used,
      $(if ($Used -eq 'да') { $What } else { '' }), $Status,
      'восстановлено из готового обзора', $Link,
      "Восстановлено 2026-05-14 из $ReviewPath; источник не найден в текущем реестре",
      'не восстановлено из готового обзора', '', 'не восстановлено из готового обзора',
      'не восстановлено из готового обзора', 'не восстановлено из готового обзора',
      'не восстановлено из готового обзора', 'не восстановлено из готового обзора',
      'не восстановлено из готового обзора', 'не восстановлено из готового обзора'
    ))
  }
  return $Rows
}

function XmlEscape([string]$Text) {
  if ($null -eq $Text) { return '' }
  return [Security.SecurityElement]::Escape($Text)
}

function New-SheetXml($Rows, [string]$FirstHeader) {
  $Headers = @($HeadersBase)
  $Headers[0] = $FirstHeader
  $AllRows = @()
  $AllRows += ,$Headers
  foreach ($R in $Rows) { $AllRows += ,$R }
  $LastRow = $AllRows.Count
  $Sb = [Text.StringBuilder]::new()
  [void]$Sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
  [void]$Sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
  [void]$Sb.AppendLine("  <dimension ref=`"A1:AB$LastRow`"/>")
  [void]$Sb.AppendLine('  <sheetViews><sheetView workbookViewId="0"/></sheetViews>')
  [void]$Sb.AppendLine('  <sheetFormatPr defaultRowHeight="15"/>')
  [void]$Sb.AppendLine('  <sheetData>')
  for ($R = 1; $R -le $AllRows.Count; $R++) {
    [void]$Sb.AppendLine("    <row r=`"$R`">")
    $Row = $AllRows[$R - 1]
    for ($C = 1; $C -le 28; $C++) {
      $Ref = (Get-ColLetters $C) + $R
      $Val = if ($C -le $Row.Count) { [string]$Row[$C - 1] } else { '' }
      [void]$Sb.AppendLine("      <c r=`"$Ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(XmlEscape $Val)</t></is></c>")
    }
    [void]$Sb.AppendLine('    </row>')
  }
  [void]$Sb.AppendLine('  </sheetData>')
  [void]$Sb.AppendLine("  <autoFilter ref=`"A1:AB$LastRow`"/>")
  [void]$Sb.AppendLine('</worksheet>')
  return $Sb.ToString()
}

function Replace-ZipText($Zip, [string]$Name, [string]$Text) {
  $Entry = Get-ZipEntryByName $Zip $Name
  if ($Entry) { $Entry.Delete() }
  $NewEntry = $Zip.CreateEntry(($Name -replace '/', '\'), [System.IO.Compression.CompressionLevel]::Optimal)
  $Writer = [IO.StreamWriter]::new($NewEntry.Open(), [Text.Encoding]::UTF8)
  try { $Writer.Write($Text) } finally { $Writer.Close() }
}

$Registry = Read-SourceRegistry

$WeekRows = [System.Collections.ArrayList]::new()
foreach ($R in (Build-StorageRows '2026-W15' 'Неделя' 'data\Архив\2026-W15\weekly_digest_multilayer.md' 'presentations\Архив\2026-W15_short_update_brief' $Registry)) { [void]$WeekRows.Add($R) }
foreach ($R in (Build-StorageRows '2026-W16' 'Неделя' 'data\Архив\2026-W16\weekly_digest_multilayer.md' 'presentations\Архив\2026-W16_director_deck' $Registry)) { [void]$WeekRows.Add($R) }
foreach ($R in (Build-StorageRows '2026-W17' 'Неделя' 'data\Архив\2026-W17\weekly_digest_multilayer.md' 'presentations\Архив\2026-W17_short_update_brief' $Registry)) { [void]$WeekRows.Add($R) }

$MonthRows = [System.Collections.ArrayList]::new()
foreach ($R in (Build-StorageRows '2026-01' 'Месяц' 'data\Архив\2026-01\monthly_digest.md' 'presentations\Архив\2026-01_monthly_brief' $Registry)) { [void]$MonthRows.Add($R) }
foreach ($R in (Build-StorageRows '2026-02' 'Месяц' 'data\Архив\2026-02\monthly_digest.md' 'presentations\Архив\2026-02_monthly_director_deck' $Registry)) { [void]$MonthRows.Add($R) }
foreach ($R in (Build-StorageRows '2026-03' 'Месяц' 'data\Архив\2026-03\monthly_digest.md' 'presentations\Архив\2026-03_monthly_director_deck' $Registry)) { [void]$MonthRows.Add($R) }
foreach ($R in (Build-StorageRows '2026-04' 'Месяц' 'data\2026-04\monthly_digest.md' 'presentations\2026-04_monthly_director_deck' $Registry)) { [void]$MonthRows.Add($R) }

$QuarterRows = [System.Collections.ArrayList]::new()
foreach ($R in (Build-StorageRows '2026-Q1' 'Квартал' 'data\2026-Q1\quarterly_digest.md' 'presentations\2026-Q1_quarterly_director_deck' $Registry)) { [void]$QuarterRows.Add($R) }

$HalfYearRows = [System.Collections.ArrayList]::new()

$YearRows = [System.Collections.ArrayList]::new()
foreach ($R in (Build-StorageRows '2025' 'Год' 'data\2025\annual_digest.md' 'presentations\2025_annual_director_deck' $Registry)) { [void]$YearRows.Add($R) }

$Zip = [System.IO.Compression.ZipFile]::Open($StoragePath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
  Replace-ZipText $Zip 'xl/worksheets/sheet3.xml' (New-SheetXml $WeekRows 'Неделя')
  Replace-ZipText $Zip 'xl/worksheets/sheet4.xml' (New-SheetXml $MonthRows 'Месяц')
  Replace-ZipText $Zip 'xl/worksheets/sheet5.xml' (New-SheetXml $QuarterRows 'Квартал')
  Replace-ZipText $Zip 'xl/worksheets/sheet6.xml' (New-SheetXml $HalfYearRows 'Полугодие')
  Replace-ZipText $Zip 'xl/worksheets/sheet7.xml' (New-SheetXml $YearRows 'Год')
} finally {
  $Zip.Dispose()
}

[PSCustomObject]@{
  RegistrySources = $Registry.Count
  WeekRows = $WeekRows.Count
  MonthRows = $MonthRows.Count
  QuarterRows = $QuarterRows.Count
  HalfYearRows = $HalfYearRows.Count
  YearRows = $YearRows.Count
  StoragePath = $StoragePath
} | Format-List
