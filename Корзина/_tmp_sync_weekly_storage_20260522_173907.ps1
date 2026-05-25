Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ErrorActionPreference = 'Stop'
$Storage = (Resolve-Path '.\Хранилище.xlsx').Path
$Sources = (Resolve-Path '.\Источники данных.xlsx').Path
$Digest = (Resolve-Path '.\data\2026-05-01_2026-05-14\weekly_digest_multilayer.md').Path
$Period = '2026-05-01_2026-05-14'
$PeriodText = '01.05.2026-14.05.2026'

function ZipEntry($zip, $name) {
  $alt = $name -replace '/', '\'
  return $zip.Entries | Where-Object { $_.FullName -eq $name -or $_.FullName -eq $alt } | Select-Object -First 1
}
function ZipRead($zip, $name) {
  $entry = ZipEntry $zip $name
  if (-not $entry) { throw "Missing entry $name" }
  $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
  try { $reader.ReadToEnd() } finally { $reader.Close() }
}
function ZipWrite($zip, $name, $text) {
  $entry = ZipEntry $zip $name
  if ($entry) { $entry.Delete() }
  $new = $zip.CreateEntry($name, [IO.Compression.CompressionLevel]::Optimal)
  $writer = [IO.StreamWriter]::new($new.Open(), [Text.UTF8Encoding]::new($false))
  try { $writer.Write($text) } finally { $writer.Close() }
}
function ColNum($ref) {
  $letters = $ref -replace '\d', ''
  $n = 0
  foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1) }
  return $n
}
function ColName($n) {
  $text = ''
  while ($n -gt 0) { $n--; $text = [char]([int](65 + ($n % 26))) + $text; $n = [int][Math]::Floor($n / 26) }
  return $text
}
function X($text) {
  if ($null -eq $text) { return '' }
  return [Security.SecurityElement]::Escape([string]$text)
}
function CellText($cell, $shared) {
  if ($cell.is) { return $cell.is.InnerText }
  $value = [string]$cell.v
  if ($cell.t -eq 's' -and $value -ne '') { return $shared[[int]$value] }
  return $value
}
function Shared($zip) {
  [xml]$sst = ZipRead $zip 'xl/sharedStrings.xml'
  $result = @()
  foreach ($si in $sst.sst.si) {
    if ($si.t) { $result += [string]$si.t } else { $result += (($si.r | ForEach-Object { [string]$_.t }) -join '') }
  }
  return ,$result
}
function Registry() {
  $zip = [IO.Compression.ZipFile]::OpenRead($Sources)
  try {
    $shared = Shared $zip
    [xml]$wb = ZipRead $zip 'xl/workbook.xml'
    [xml]$rels = ZipRead $zip 'xl/_rels/workbook.xml.rels'
    $targets = @{}
    foreach ($rel in $rels.Relationships.Relationship) { $targets[$rel.Id] = $rel.Target }
    $skip = @('Доступность_источников', 'Частично_доступные', 'детали проблем')
    $rows = @()
    foreach ($sheet in $wb.workbook.sheets.sheet) {
      if ($skip -contains [string]$sheet.name) { continue }
      $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
      [xml]$xml = ZipRead $zip ('xl/' + $targets[$rid])
      $data = @($xml.worksheet.sheetData.row)
      $headers = @{}
      foreach ($cell in $data[0].c) { $headers[(CellText $cell $shared)] = ColNum $cell.r }
      $nameCol = if ($headers.ContainsKey('Название')) { $headers['Название'] } elseif ($headers.ContainsKey('название источника')) { $headers['название источника'] } else { 1 }
      $collectCol = if ($headers.ContainsKey('Что собираем')) { $headers['Что собираем'] } elseif ($headers.ContainsKey('что собираем')) { $headers['что собираем'] } else { $nameCol + 1 }
      $urlCol = if ($headers.ContainsKey('ссылка')) { $headers['ссылка'] } else { $nameCol + 2 }
      foreach ($row in $data | Select-Object -Skip 1) {
        $values = @{}
        foreach ($cell in $row.c) { $values[(ColNum $cell.r)] = CellText $cell $shared }
        $name = ([string]$values[$nameCol]).Trim()
        if ($name) {
          $rows += [pscustomobject]@{ Source = $name; Type = [string]$sheet.name; Collect = ([string]$values[$collectCol]).Trim(); Url = ([string]$values[$urlCol]).Trim() }
        }
      }
    }
    return $rows | Sort-Object Source -Unique
  } finally { $zip.Dispose() }
}
function Decisions() {
  $lines = Get-Content -LiteralPath $Digest -Encoding UTF8
  $result = @{}
  $start = [Array]::FindIndex($lines, [Predicate[string]]{ param($line) $line -match '^\| Источник \|' })
  for ($i = $start + 2; $i -lt $lines.Count -and $lines[$i].StartsWith('|'); $i++) {
    $cells = @($lines[$i].Trim('|').Split('|') | ForEach-Object { $_.Trim() })
    $result[$cells[0]] = [pscustomobject]@{ Status = $cells[1]; What = $cells[2]; Used = $cells[3] }
  }
  return $result
}
function Evidence($source) {
  if ($source.Type -eq 'макро') { return 'официальный источник' }
  if ($source.Type -eq 'платформы') { return 'платформа с данными' }
  if ($source.Type -eq 'сми') { return 'деловое медиа' }
  if ($source.Type -eq 'тг') { return 'ручная проверка / слабый сигнал' }
  if ($source.Type -match 'миров|глобаль') { return 'внешний контур' }
  return 'профильный источник'
}
function Row($n, $values) {
  $sb = [Text.StringBuilder]::new()
  [void]$sb.Append("<row r=`"$n`">")
  for ($i = 0; $i -lt $values.Count; $i++) {
    [void]$sb.Append("<c r=`"$(ColName ($i + 1))$n`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(X $values[$i])</t></is></c>")
  }
  [void]$sb.Append('</row>')
  return $sb.ToString()
}

$headers = @('Неделя','Название источника','дата сбора информации','тип источника: сайт / telegram / иное','заголовок конкретного материала / статус отсутствия данных','что собираем','тема предмета','регион','отрасль','тип сигнала','тональность','важность','контур','использован в обзоре','причина включения','статус проверки источника','уровень доказательности','ссылка','комментарий','стратегический сигнал','объект стратегического сигнала','влияет на рынок труда','влияет на ИТ-рынок','влияет на значимых работодателей','влияет на стратегических клиентов','сдвиг в логике отрасли','механизм влияния','шум или общий сдвиг')
$registry = Registry | Where-Object { $_.Source -ne 'System.Xml.XmlElement' }
$decisions = Decisions
$registryByName = @{}
foreach ($source in $registry) { $registryByName[$source.Source] = $source }
$body = [Text.StringBuilder]::new()
[void]$body.Append((Row 1 $headers))
$rowNum = 2
foreach ($sourceName in ($decisions.Keys | Sort-Object)) {
  $d = $decisions[$sourceName]
  $source = if ($registryByName.ContainsKey($sourceName)) { $registryByName[$sourceName] } else { [pscustomobject]@{ Source = $sourceName; Type = 'реестр источников'; Collect = ''; Url = '' } }
  $reason = if ($d.Used -eq 'да') { $d.What } else { '' }
  $strategic = if ($source.Source -eq 'AP / Business and Finance') { @('да','мировой технологический сектор и ИИ','да','да','да','нет','да','ИИ и давление на эффективность -> перераспределение ролей и бюджетов -> более строгий найм','общий сдвиг') } else { @('','','','','','','','','') }
  $values = @($Period,$source.Source,'2026-05-22',$source.Type,$d.What,$source.Collect,'рынок труда / ИТ / стратегические сигналы',$(if ($source.Type -match 'миров|глобаль') {'внешний контур'} else {'Россия'}),'рынок труда; ИТ','проверка источника','нейтральная',$(if ($d.Used -eq 'да') {'высокая'} else {'служебная'}),$source.Type,$d.Used,$reason,$d.Status,(Evidence $source),$source.Url,"Проверено для недельного среза $PeriodText. Если профильный проход в среде ненадежен, статус оставлен как ручная проверка.") + $strategic
  [void]$body.Append((Row $rowNum $values))
  $rowNum++
}
$sheet = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><dimension ref=`"A1:AB$($rowNum - 1)`"/><sheetViews><sheetView workbookViewId=`"0`"/></sheetViews><sheetFormatPr defaultRowHeight=`"15`"/><sheetData>$body</sheetData></worksheet>"

$zip = [IO.Compression.ZipFile]::Open($Storage, [IO.Compression.ZipArchiveMode]::Update)
try {
  ZipWrite $zip 'xl/worksheets/sheet3.xml' $sheet
  [xml]$journal = ZipRead $zip 'xl/worksheets/sheet8.xml'
  if (-not $journal.OuterXml.Contains('Недельный срез 01.05.2026-14.05.2026')) {
    $rows = @($journal.DocumentElement.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']"))
    $next = ([int](($rows | ForEach-Object { $_.GetAttribute('r') } | Measure-Object -Maximum).Maximum)) + 1
    $event = @('2026-05-14','Мировой технологический сектор и ИИ','крупные увольнения в мировой технологии','AP зафиксировал, что объявления о сокращениях в технологическом секторе все чаще связываются с ИИ, перераспределением ресурсов и эффективностью.','AP / Business and Finance','https://apnews.com/article/65f9944fa25306bf5c975dd94805731e','ИТ-рынок, рынок труда, значимые работодатели, логика отрасли','сигнал замещения трудом ИИ','ИИ и давление на эффективность -> перераспределение ролей и бюджетов -> более строгий найм','общий сдвиг','да','Недельный срез 01.05.2026-14.05.2026; внешний сигнал, не прямое доказательство по России')
    $fragment = [xml]("<root xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`">$(Row $next $event)</root>")
    $journal.worksheet.sheetData.AppendChild($journal.ImportNode($fragment.root.row, $true)) | Out-Null
    $journal.worksheet.dimension.ref = "A1:L$next"
    ZipWrite $zip 'xl/worksheets/sheet8.xml' $journal.OuterXml
  }
} finally { $zip.Dispose() }
Write-Output ('REGISTRY=' + $decisions.Count)
Write-Output ('WEEKLY_ROWS=' + ($rowNum - 2))
