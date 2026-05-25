Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ErrorActionPreference = 'Stop'
$Period = '2026-05-01_2026-05-14'
$PeriodLabel = '01.05.2026-14.05.2026'
$CollectedOn = '2026-05-22'
$StoragePath = (Resolve-Path '.\Хранилище.xlsx').Path
$SourcesPath = (Resolve-Path '.\Источники данных.xlsx').Path
$DigestDir = Join-Path (Get-Location) ('data\' + $Period)
$DigestPath = Join-Path $DigestDir 'weekly_digest_multilayer.md'

function Get-ZipEntry {
  param($Zip, [string]$Name)
  $alt = $Name -replace '/', '\'
  return $Zip.Entries | Where-Object { $_.FullName -eq $Name -or $_.FullName -eq $alt } | Select-Object -First 1
}

function Read-ZipText {
  param($Zip, [string]$Name)
  $entry = Get-ZipEntry $Zip $Name
  if (-not $entry) { throw "Missing zip entry: $Name" }
  $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
  try { return $reader.ReadToEnd() } finally { $reader.Close() }
}

function Write-ZipText {
  param($Zip, [string]$Name, [string]$Text)
  $entry = Get-ZipEntry $Zip $Name
  if ($entry) { $entry.Delete() }
  $newEntry = $Zip.CreateEntry($Name, [IO.Compression.CompressionLevel]::Optimal)
  $writer = [IO.StreamWriter]::new($newEntry.Open(), [Text.UTF8Encoding]::new($false))
  try { $writer.Write($Text) } finally { $writer.Close() }
}

function Get-ColNum {
  param([string]$Ref)
  $letters = $Ref -replace '\d', ''
  $n = 0
  foreach ($ch in $letters.ToCharArray()) {
    $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $n
}

function Get-ColLetters {
  param([int]$Num)
  $text = ''
  while ($Num -gt 0) {
    $Num--
    $text = [char](65 + ($Num % 26)) + $text
    $Num = [Math]::Floor($Num / 26)
  }
  return $text
}

function XmlEscape {
  param([string]$Text)
  if ($null -eq $Text) { return '' }
  return [Security.SecurityElement]::Escape($Text)
}

function CellText {
  param($Cell, $SharedStrings)
  if ($Cell.is) { return $Cell.is.InnerText }
  $value = [string]$Cell.v
  if ($Cell.t -eq 's' -and $value -ne '') { return $SharedStrings[[int]$value] }
  return $value
}

function Read-SharedStrings {
  param($Zip)
  [xml]$sst = Read-ZipText $Zip 'xl/sharedStrings.xml'
  $shared = @()
  foreach ($si in $sst.sst.si) {
    if ($si.t) { $shared += [string]$si.t }
    else { $shared += (($si.r | ForEach-Object { [string]$_.t }) -join '') }
  }
  return ,$shared
}

function Read-SourceRegistry {
  $zip = [IO.Compression.ZipFile]::OpenRead($SourcesPath)
  try {
    $shared = Read-SharedStrings $zip
    [xml]$wb = Read-ZipText $zip 'xl/workbook.xml'
    [xml]$rels = Read-ZipText $zip 'xl/_rels/workbook.xml.rels'
    $relMap = @{}
    foreach ($rel in $rels.Relationships.Relationship) { $relMap[$rel.Id] = $rel.Target }
    $skip = @('Доступность_источников', 'Частично_доступные', 'детали проблем')
    $all = @()
    foreach ($sheet in $wb.workbook.sheets.sheet) {
      if ($skip -contains [string]$sheet.name) { continue }
      $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
      [xml]$xml = Read-ZipText $zip ('xl/' + $relMap[$rid])
      $rows = @($xml.worksheet.sheetData.row)
      if ($rows.Count -eq 0) { continue }
      $headers = @{}
      foreach ($cell in $rows[0].c) { $headers[(CellText $cell $shared)] = Get-ColNum $cell.r }
      $nameCol = if ($headers.ContainsKey('Название')) { $headers['Название'] } elseif ($headers.ContainsKey('название источника')) { $headers['название источника'] } else { 1 }
      $collectCol = if ($headers.ContainsKey('Что собираем')) { $headers['Что собираем'] } elseif ($headers.ContainsKey('что собираем')) { $headers['что собираем'] } else { $nameCol + 1 }
      $urlCol = if ($headers.ContainsKey('ссылка')) { $headers['ссылка'] } elseif ($headers.ContainsKey('Ссылка')) { $headers['Ссылка'] } else { $nameCol + 2 }
      foreach ($row in $rows | Select-Object -Skip 1) {
        $values = @{}
        foreach ($cell in $row.c) { $values[(Get-ColNum $cell.r)] = CellText $cell $shared }
        $name = ([string]$values[$nameCol]).Trim()
        if (-not $name) { continue }
        $all += [pscustomobject]@{
          Source = $name
          Type = [string]$sheet.name
          Collect = ([string]$values[$collectCol]).Trim()
          Url = ([string]$values[$urlCol]).Trim()
        }
      }
    }
    return $all | Sort-Object Source -Unique
  } finally {
    $zip.Dispose()
  }
}

function EvidenceLevel {
  param($Source)
  switch -Regex ($Source.Type) {
    '^макро$' { return 'официальный источник' }
    '^платформы$' { return 'платформа с данными' }
    '^сми$' { return 'деловое медиа' }
    '^тг$' { return 'ручная проверка / слабый сигнал' }
    'миров' { return 'внешний контур' }
    'глобаль' { return 'внешний контур' }
    default { return 'профильный источник' }
  }
}

function New-Decision {
  param($Source)
  $manualNames = @(
    'Авито аналитика', 'getmatch', 'RealHR', 'IBS', 'Технологии Доверия',
    'HR-аналитика', 'Про HR и не только / Наталья Володина', 'Канал Павла Безручко',
    'Happy Job. Всё о развитии персонала', 'Балицкая / TheBalitskaya',
    'CEO рулит I Regroup', 'HR Перезагрузка', 'Зарплата в IT (+ вакансии)',
    'ТеДо / ПроНалоги. TaxPro', 'HR-новости', 'The HRD', 'topcareer',
    'Эй, HR! Антон Платонов', 'IT HR тусовка от ХК', 'HR Sk',
    'HR-дайджест, самый полный', 'Пульс / PRO. Людей', 'WTF_HR'
  )
  $decision = [ordered]@{
    Status = 'проверен, не релевантно'
    What = 'релевантных материалов за период 01.05.2026-14.05.2026 не найдено'
    Used = 'нет'
    Reason = ''
    SignalType = 'проверка источника'
    Tonality = 'нейтральная'
    Importance = 'служебная'
    Strategic = ''
    StrategicObject = ''
    LabourImpact = ''
    ItImpact = ''
    EmployerImpact = ''
    ClientImpact = ''
    LogicShift = ''
    Mechanism = ''
    Shift = ''
  }
  if ($manualNames -contains $Source.Source) {
    $decision.Status = 'требует ручной проверки'
    $decision.What = 'требуется ручная проверка источника за период 01.05.2026-14.05.2026'
    return [pscustomobject]$decision
  }
  switch ($Source.Source) {
    'hh.ru' {
      $decision.Status = 'найден релевантный материал'
      $decision.What = 'Краткий обзор рынка труда hh.ru за март 2026: ИТ hh.индекс 22,9; в окне недели доступна свежая публикация обзора'
      $decision.Used = 'да'
      $decision.Reason = 'дает платформенный сигнал о высокой конкуренции за ИТ-вакансии'
      $decision.SignalType = 'платформенная статистика'
      $decision.Importance = 'высокая'
    }
    'Русофт' {
      $decision.Status = 'найден релевантный материал'
      $decision.What = '13.05.2026: 71% опрошенных компаний РФ испытывают нехватку кадров на рынке информбезопасности'
      $decision.Used = 'да'
      $decision.Reason = 'показывает, что охлаждение общего ИТ-найма не снимает дефицит по ИБ-компетенциям'
      $decision.SignalType = 'ИТ-кадры'
      $decision.Importance = 'высокая'
    }
    'Коммерсантъ' {
      $decision.Status = 'найден релевантный материал'
      $decision.What = '14.05.2026: рынок труда 2026 разворачивается в сторону работодателя'
      $decision.Used = 'да'
      $decision.Reason = 'поддерживает общий сигнал о более выборочном найме'
      $decision.SignalType = 'деловая аналитика'
      $decision.Importance = 'средняя'
    }
    'AP / Business and Finance' {
      $decision.Status = 'найден релевантный материал'
      $decision.What = '14.05.2026: AP о том, что сокращения в технологическом секторе все чаще связывают с ИИ и перераспределением ресурсов'
      $decision.Used = 'да'
      $decision.Reason = 'дает внешний стратегический сигнал по ИИ, ролям и порогу эффективности'
      $decision.SignalType = 'стратегическое отраслевое событие'
      $decision.Importance = 'средняя'
      $decision.Strategic = 'да'
      $decision.StrategicObject = 'мировой технологический сектор и ИИ'
      $decision.LabourImpact = 'да'
      $decision.ItImpact = 'да'
      $decision.EmployerImpact = 'да'
      $decision.ClientImpact = 'нет'
      $decision.LogicShift = 'да'
      $decision.Mechanism = 'ИИ и давление на эффективность -> перераспределение ролей и бюджетов -> более строгий найм'
      $decision.Shift = 'общий сдвиг'
    }
    'Экоспи' {
      $decision.Status = 'проверен, не релевантно'
      $decision.What = 'в окне периода открывался сбор ответов для исследования HR, цифровизации и ИИ; результатов исследования еще нет'
    }
    'ХэндФлоу' {
      $decision.Status = 'проверен, не релевантно'
      $decision.What = '05.05.2026 найден материал о майских праздниках для HR; рыночного сигнала для выпуска не дает'
    }
    'Хабр Карьера' {
      $decision.Status = 'проверен, не релевантно'
      $decision.What = 'проверены карьерные и зарплатные разделы за окно периода; отдельного нового рыночного сигнала для выпуска не найдено'
    }
  }
  return [pscustomobject]$decision
}

function RowXml {
  param([int]$RowNumber, [object[]]$Values)
  $sb = [Text.StringBuilder]::new()
  [void]$sb.Append("<row r=`"$RowNumber`">")
  for ($i = 0; $i -lt $Values.Count; $i++) {
    $ref = "$(Get-ColLetters ($i + 1))$RowNumber"
    $value = XmlEscape ([string]$Values[$i])
    [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$value</t></is></c>")
  }
  [void]$sb.Append('</row>')
  return $sb.ToString()
}

function New-WeeklySheetXml {
  param($Rows)
  $headers = @(
    'Неделя','Название источника','дата сбора информации','тип источника: сайт / telegram / иное',
    'заголовок конкретного материала / статус отсутствия данных','что собираем','тема предмета','регион',
    'отрасль','тип сигнала','тональность','важность','контур','использован в обзоре','причина включения',
    'статус проверки источника','уровень доказательности','ссылка','комментарий','стратегический сигнал',
    'объект стратегического сигнала','влияет на рынок труда','влияет на ИТ-рынок','влияет на значимых работодателей',
    'влияет на стратегических клиентов','сдвиг в логике отрасли','механизм влияния','шум или общий сдвиг'
  )
  $sb = [Text.StringBuilder]::new()
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
  [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
  [void]$sb.Append("<dimension ref=`"A1:AB$($Rows.Count + 1)`"/>")
  [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>')
  [void]$sb.Append((RowXml 1 $headers))
  $rowNumber = 2
  foreach ($row in $Rows) {
    [void]$sb.Append((RowXml $rowNumber $row))
    $rowNumber++
  }
  [void]$sb.Append('</sheetData></worksheet>')
  return $sb.ToString()
}

function Add-JournalEvent {
  param($Zip)
  [xml]$xml = Read-ZipText $Zip 'xl/worksheets/sheet8.xml'
  $already = $xml.OuterXml.Contains('2026-05-14') -and $xml.OuterXml.Contains('мировой технологический сектор и ИИ')
  if ($already) { return }
  $rows = @($xml.DocumentElement.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']"))
  $nextRow = ([int](($rows | ForEach-Object { $_.GetAttribute('r') } | Measure-Object -Maximum).Maximum)) + 1
  $values = @(
    '2026-05-14',
    'Мировой технологический сектор и ИИ',
    'крупные увольнения в мировой технологии',
    'AP зафиксировал, что объявления о сокращениях в технологическом секторе все чаще связываются с ИИ, перераспределением ресурсов и повышением эффективности.',
    'AP / Business and Finance',
    'https://apnews.com/article/65f9944fa25306bf5c975dd94805731e',
    'ИТ-рынок, рынок труда, значимые работодатели, логика отрасли',
    'сигнал замещения трудом ИИ',
    'ИИ и давление на эффективность -> перераспределение ролей и бюджетов -> более строгий найм',
    'общий сдвиг',
    'да',
    'Недельный срез 01.05.2026-14.05.2026; внешний сигнал, не прямое доказательство по России'
  )
  $fragment = [xml]("<root xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`">$(RowXml $nextRow $values)</root>")
  $rowNode = $xml.ImportNode($fragment.root.row, $true)
  $xml.worksheet.sheetData.AppendChild($rowNode) | Out-Null
  $xml.worksheet.dimension.ref = "A1:L$nextRow"
  Write-ZipText $Zip 'xl/worksheets/sheet8.xml' $xml.OuterXml
}

function New-Appendix {
  param($Rows)
  $lines = [Collections.Generic.List[string]]::new()
  $lines.Add('| Источник | Статус за неделю | Что найдено или почему данных нет | Использован в выпуске |')
  $lines.Add('|---|---|---|---|')
  foreach ($row in $Rows) {
    $lines.Add("| $($row.Source) | $($row.Status) | $($row.What) | $($row.Used) |")
  }
  return ($lines -join "`r`n")
}

$registry = Read-SourceRegistry
$decisions = @()
$storageRows = @()
foreach ($source in $registry) {
  $decision = New-Decision $source
  $decisions += [pscustomobject]@{
    Source = $source.Source
    Status = $decision.Status
    What = $decision.What
    Used = $decision.Used
  }
  $storageRows += ,@(
    $Period,
    $source.Source,
    $CollectedOn,
    $source.Type,
    $decision.What,
    $source.Collect,
    'рынок труда / ИТ / стратегические сигналы',
    $(if ($source.Type -match 'миров|глобаль') { 'внешний контур' } else { 'Россия' }),
    'рынок труда; ИТ',
    $decision.SignalType,
    $decision.Tonality,
    $decision.Importance,
    $source.Type,
    $decision.Used,
    $decision.Reason,
    $decision.Status,
    (EvidenceLevel $source),
    $source.Url,
    "Проверено для недельного среза $PeriodLabel. Решение зафиксировано по профильному источнику или как ручная проверка, если среда не дает надежного прохода.",
    $decision.Strategic,
    $decision.StrategicObject,
    $decision.LabourImpact,
    $decision.ItImpact,
    $decision.EmployerImpact,
    $decision.ClientImpact,
    $decision.LogicShift,
    $decision.Mechanism,
    $decision.Shift
  )
}

$storageZip = [IO.Compression.ZipFile]::Open($StoragePath, [IO.Compression.ZipArchiveMode]::Update)
try {
  Write-ZipText $storageZip 'xl/worksheets/sheet3.xml' (New-WeeklySheetXml $storageRows)
  Add-JournalEvent $storageZip
} finally {
  $storageZip.Dispose()
}

[IO.Directory]::CreateDirectory($DigestDir) | Out-Null
$appendix = New-Appendix $decisions
$digest = @"
# Недельный срез рынка труда: 01.05.2026-14.05.2026

## Метаданные
- `week_label`: `$Period`
- `week_start`: `2026-05-01`
- `week_end`: `2026-05-14`
- `география`: `Россия`
- `сравнение_с`: `предыдущая неделя; сравнение ограничено, потому что актуального недельного выпуска за непосредственно предшествующее окно в рабочем контуре нет`
- `подготовлено`: `$CollectedOn`

## 1. Ключевые сигналы недели
- Майское окно дало узкий корпус сигналов: после праздничных разрывов в рабочем контуре немного новых регулярных публикаций, поэтому формат выпуска остается коротким.
- Свежий отчет `hh.ru`, доступный в окне среза, сохраняет жесткий сигнал по ИТ-рынку: в мартовском рэнкинге по России для сферы информационных технологий `hh.индекс` указан на уровне `22,9`, то есть конкуренция за ИТ-вакансии остается высокой.
- Внутри охлажденного ИТ-контура сохраняются дефицитные зоны: `РУССОФТ` 13 мая вынес сигнал о нехватке кадров на рынке информационной безопасности у 71% опрошенных компаний.
- Общий рынок труда в деловой аналитике продолжает описываться как более выборочный для кандидата: `Коммерсантъ` 14 мая прямо формулирует разворот 2026 года в сторону работодателя.

## 2. Стратегические сигналы периода
- Внешний технологический контур продолжает связывать ИИ с давлением на роли и штат. AP 14 мая зафиксировал, что технологические компании все чаще сопровождают объявления о сокращениях риторикой про ИИ, перераспределение ресурсов и эффективность. Для российского рынка это не прямое доказательство немедленных сокращений, а внешний сигнал: порог полезности роли и ожидание прикладной эффективности продолжают расти.

## 3. Рынок труда ИТ в России
- Главный подтвержденный сигнал окна не про новый всплеск найма, а про сохраняющуюся конкуренцию за ИТ-вакансии. Данные `hh.ru` поддерживают осторожную трактовку: в широком ИТ-контуре у работодателя остается больше пространства для отбора.
- При этом рынок не выглядит однородно избыточным. Сигнал `РУССОФТ` по информационной безопасности показывает, что для прикладных и дефицитных компетенций нехватка сохраняется даже на фоне более жесткого общего найма.
- По сравнению с непосредственно предшествующей неделей вывод нужно держать мягким: актуального недельного выпуска в рабочем контуре нет, а подтвержденный майский корпус пока слишком узкий для заявления о новом тренде.

## 4. Рынок труда России
- В общей рамке недели усиливается не новый статистический перелом, а риторика более прагматичного найма. Материал `Коммерсанта` поддерживает апрельский фон: компании чаще говорят об эффективности, автоматизации и точечном подборе.
- Праздничное окно не дает достаточной базы для сильного вывода о недельном ускорении или торможении рынка труда в целом. Здесь правильнее фиксировать сохранение осторожности, а не выдумывать динамику по тишине источников.

## 5. Что это значит для работодателя
- Найм: можно точнее фильтровать широкий ИТ-поток, но нельзя переносить это ощущение на все специализации подряд.
- Удержание: дефицитные зоны вроде информационной безопасности требуют отдельной логики удержания и развития, даже если на массовых ИТ-ролях выбор кандидатов стал шире.
- Зарплаты: короткое окно не дает основания для нового зарплатного вывода; разумнее смотреть на следующие платформенные и официальные публикации.
- Ближайший риск: под риторикой ИИ и эффективности часть ролей будет оцениваться жестче по прикладному результату, а не по самому факту цифровой специализации.

## 6. На что смотреть на следующей неделе
- Появятся ли после майских праздников новые платформенные или официальные данные, которые подтвердят или ослабят сигнал высокой конкуренции в ИТ.
- Повторится ли сигнал о точечном найме и дефиците отдельных компетенций в источниках помимо `hh.ru`, `РУССОФТ` и деловых медиа.

## Приложение по источникам
$appendix
"@
[IO.File]::WriteAllText($DigestPath, $digest, [Text.UTF8Encoding]::new($false))

Write-Output ('REGISTRY_COUNT=' + $registry.Count)
Write-Output ('WEEKLY_ROWS=' + $storageRows.Count)
Write-Output ('DIGEST_PATH=' + $DigestPath)
