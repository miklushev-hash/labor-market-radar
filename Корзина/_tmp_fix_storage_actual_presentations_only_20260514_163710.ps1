Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

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
  if ($Cell.is) { return $Cell.is.InnerText }
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
  $N = [System.Web.HttpUtility]::HtmlDecode($N)
  $N = $N -replace '\s+', ' '
  $N = $N.Trim()
  $N = $N -replace '\s*[:：].*$', ''
  return $N.Trim().ToLowerInvariant()
}

function Convert-HtmlCellToText([string]$Html) {
  $T = $Html -replace '<br\s*/?>', '; '
  $T = $T -replace '<[^>]+>', ''
  $T = [System.Web.HttpUtility]::HtmlDecode($T)
  $T = $T -replace '\s+', ' '
  return $T.Trim()
}

function Parse-HtmlSourceRows([string]$Path) {
  $Rows = [System.Collections.ArrayList]::new()
  if (-not (Test-Path -LiteralPath $Path)) { return $Rows }
  $Html = Get-Content -LiteralPath $Path -Encoding UTF8 -Raw
  foreach ($TableMatch in [regex]::Matches($Html, '<table\b.*?</table>', 'Singleline,IgnoreCase')) {
    $Table = $TableMatch.Value
    $HeaderCells = @()
    $HeaderMatch = [regex]::Match($Table, '<tr\b.*?</tr>', 'Singleline,IgnoreCase')
    if (-not $HeaderMatch.Success) { continue }
    foreach ($CellMatch in [regex]::Matches($HeaderMatch.Value, '<th\b[^>]*>(.*?)</th>', 'Singleline,IgnoreCase')) {
      $HeaderCells += (Convert-HtmlCellToText $CellMatch.Groups[1].Value)
    }
    if ($HeaderCells.Count -eq 0 -or $HeaderCells[0] -notmatch 'Источник') { continue }
    $StatusIdx = [Array]::FindIndex($HeaderCells, [Predicate[string]]{ param($H) $H -match 'Статус' })
    $WhatIdx = [Array]::FindIndex($HeaderCells, [Predicate[string]]{ param($H) $H -match 'Что' })
    $UsedIdx = [Array]::FindIndex($HeaderCells, [Predicate[string]]{ param($H) $H -match 'Использован' })
    foreach ($RowMatch in [regex]::Matches($Table, '<tr\b.*?</tr>', 'Singleline,IgnoreCase') | Select-Object -Skip 1) {
      $Cells = @()
      foreach ($CellMatch in [regex]::Matches($RowMatch.Value, '<td\b[^>]*>(.*?)</td>', 'Singleline,IgnoreCase')) {
        $Cells += (Convert-HtmlCellToText $CellMatch.Groups[1].Value)
      }
      if ($Cells.Count -eq 0) { continue }
      $Source = $Cells[0]
      if (-not $Source) { continue }
      $Status = if ($StatusIdx -ge 0 -and $Cells.Count -gt $StatusIdx) { $Cells[$StatusIdx] } else { 'использовано в выпуске' }
      $What = if ($WhatIdx -ge 0 -and $Cells.Count -gt $WhatIdx) { $Cells[$WhatIdx] } else { '' }
      $Used = if ($UsedIdx -ge 0 -and $Cells.Count -gt $UsedIdx) { $Cells[$UsedIdx] } else { 'да' }
      if ($Status -match 'не использовано' -or $Used -match 'нет') { $Used = 'нет' }
      elseif ($Used -notmatch 'нет') { $Used = 'да' }
      [void]$Rows.Add([PSCustomObject]@{
        Source = $Source
        Key = Normalize-SourceName $Source
        Status = $Status
        What = $What
        Used = $Used
        Url = ''
      })
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
      foreach ($C in $Rows[0].c) { $HeaderByName[(Get-CellText $C $Shared)] = Get-ColNum $C.r }
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

function Merge-Rows($Rows) {
  $Map = @{}
  foreach ($R in $Rows) {
    if (-not $R.Key) { continue }
    if (-not $Map.ContainsKey($R.Key)) {
      $Map[$R.Key] = [PSCustomObject]@{
        Source = $R.Source
        Statuses = [System.Collections.Generic.List[string]]::new()
        Whats = [System.Collections.Generic.List[string]]::new()
        UsedValues = [System.Collections.Generic.List[string]]::new()
      }
    }
    if ($R.Status) { $Map[$R.Key].Statuses.Add($R.Status) }
    if ($R.What) { $Map[$R.Key].Whats.Add($R.What) }
    if ($R.Used) { $Map[$R.Key].UsedValues.Add($R.Used) }
  }
  return $Map
}

function Build-StorageRows($Period, $HtmlPath, $PdfPath, $Registry) {
  $Parsed = Merge-Rows (Parse-HtmlSourceRows $HtmlPath)
  $Rows = [System.Collections.ArrayList]::new()
  $Seen = @{}
  foreach ($Src in $Registry) {
    $ParsedRow = $null
    if ($Parsed.ContainsKey($Src.Key)) { $ParsedRow = $Parsed[$Src.Key] }
    $Status = 'не восстановлено из актуальной презентации; требуется ручная проверка источника'
    $What = 'статус по этому источнику отсутствует в актуальной презентации'
    $Used = 'нет'
    $Reason = ''
    $Evidence = 'не восстановлено'
    if ($ParsedRow) {
      $Status = (($ParsedRow.Statuses | Select-Object -Unique) -join '; ')
      $What = (($ParsedRow.Whats | Select-Object -Unique) -join '; ')
      if (-not $What) { $What = 'источник указан в приложении актуальной презентации' }
      $Used = if (($ParsedRow.UsedValues | Where-Object { $_ -eq 'да' }).Count -gt 0) { 'да' } else { 'нет' }
      $Reason = if ($Used -eq 'да') { $What } else { '' }
      $Evidence = 'восстановлено из актуальной презентации'
      $Seen[$Src.Key] = $true
    }
    [void]$Rows.Add([object[]]@(
      $Period, $Src.Source, $Today, $Src.Type, $What, $Src.Collect,
      'рынок труда / ИТ / экономика', 'Россия', 'рынок труда; ИТ',
      'не восстановлено из актуальной презентации', 'не восстановлено из актуальной презентации',
      'не восстановлено из актуальной презентации', $Src.Type, $Used, $Reason,
      $Status, $Evidence, $Src.Url,
      "Восстановлено 2026-05-14 только из актуальной версии: $HtmlPath / $PdfPath",
      'не восстановлено из актуальной презентации', '',
      'не восстановлено из актуальной презентации', 'не восстановлено из актуальной презентации',
      'не восстановлено из актуальной презентации', 'не восстановлено из актуальной презентации',
      'не восстановлено из актуальной презентации', 'не восстановлено из актуальной презентации',
      'не восстановлено из актуальной презентации'
    ))
  }
  foreach ($Key in $Parsed.Keys) {
    if ($Seen.ContainsKey($Key)) { continue }
    $P = $Parsed[$Key]
    $What = (($P.Whats | Select-Object -Unique) -join '; ')
    $Status = (($P.Statuses | Select-Object -Unique) -join '; ')
    $Used = if (($P.UsedValues | Where-Object { $_ -eq 'да' }).Count -gt 0) { 'да' } else { 'нет' }
    [void]$Rows.Add([object[]]@(
      $Period, $P.Source, $Today, 'источник вне текущего реестра', $What,
      '', 'рынок труда / ИТ / экономика', 'Россия', 'рынок труда; ИТ',
      'не восстановлено из актуальной презентации', 'не восстановлено из актуальной презентации',
      'не восстановлено из актуальной презентации', 'вне текущего реестра', $Used,
      $(if ($Used -eq 'да') { $What } else { '' }), $Status,
      'восстановлено из актуальной презентации', '',
      "Восстановлено 2026-05-14 только из актуальной версии: $HtmlPath / $PdfPath; источник не найден в текущем реестре",
      'не восстановлено из актуальной презентации', '', 'не восстановлено из актуальной презентации',
      'не восстановлено из актуальной презентации', 'не восстановлено из актуальной презентации',
      'не восстановлено из актуальной презентации', 'не восстановлено из актуальной презентации',
      'не восстановлено из актуальной презентации', 'не восстановлено из актуальной презентации'
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
  $AllRows = [System.Collections.ArrayList]::new()
  [void]$AllRows.Add([object[]]$Headers)
  foreach ($R in $Rows) { [void]$AllRows.Add($R) }
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
$EmptyRows = [System.Collections.ArrayList]::new()
$MonthRows = Build-StorageRows '2026-04' 'presentations\2026-04_monthly_director_deck.html' 'presentations\2026-04_monthly_director_deck.pdf' $Registry
$QuarterRows = Build-StorageRows '2026-Q1' 'presentations\2026-Q1_quarterly_director_deck.html' 'presentations\2026-Q1_quarterly_director_deck.pdf' $Registry
$YearRows = Build-StorageRows '2025' 'presentations\2025_annual_director_deck.html' 'presentations\2025_annual_director_deck.pdf' $Registry

$Zip = [System.IO.Compression.ZipFile]::Open($StoragePath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
  Replace-ZipText $Zip 'xl/worksheets/sheet3.xml' (New-SheetXml $EmptyRows 'Неделя')
  Replace-ZipText $Zip 'xl/worksheets/sheet4.xml' (New-SheetXml $MonthRows 'Месяц')
  Replace-ZipText $Zip 'xl/worksheets/sheet5.xml' (New-SheetXml $QuarterRows 'Квартал')
  Replace-ZipText $Zip 'xl/worksheets/sheet6.xml' (New-SheetXml $EmptyRows 'Полугодие')
  Replace-ZipText $Zip 'xl/worksheets/sheet7.xml' (New-SheetXml $YearRows 'Год')
} finally {
  $Zip.Dispose()
}

[PSCustomObject]@{
  RegistrySources = $Registry.Count
  WeekRows = 0
  MonthRows = $MonthRows.Count
  QuarterRows = $QuarterRows.Count
  HalfYearRows = 0
  YearRows = $YearRows.Count
  SourceScope = 'presentations root only'
  StoragePath = $StoragePath
} | Format-List
