Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ErrorActionPreference = 'Stop'

function Get-ZipEntry {
  param($Zip, [string]$Name)
  $alt = $Name -replace '/', '\'
  return $Zip.Entries | Where-Object { $_.FullName -eq $Name -or $_.FullName -eq $alt } | Select-Object -First 1
}

function Read-EntryText {
  param($Zip, [string]$Name)
  $entry = Get-ZipEntry $Zip $Name
  if (-not $entry) { throw "Missing entry: $Name" }
  $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
  try { return $reader.ReadToEnd() } finally { $reader.Close() }
}

function Write-EntryText {
  param($Zip, [string]$Name, [string]$Text)
  $entry = Get-ZipEntry $Zip $Name
  if ($entry) { $entry.Delete() }
  $newEntry = $Zip.CreateEntry($Name, [IO.Compression.CompressionLevel]::Optimal)
  $writer = [IO.StreamWriter]::new($newEntry.Open(), [Text.UTF8Encoding]::new($false))
  try { $writer.Write($Text) } finally { $writer.Close() }
}

function Split-MarkdownRow {
  param([string]$Line)
  $text = $Line.Trim()
  if ($text.StartsWith('|')) { $text = $text.Substring(1) }
  if ($text.EndsWith('|')) { $text = $text.Substring(0, $text.Length - 1) }
  return @($text -split '\|' | ForEach-Object { $_.Trim() })
}

function Get-AprilSourceMap {
  $lines = Get-Content -LiteralPath '.\data\2026-04\monthly_digest.md' -Encoding UTF8
  $map = @{}
  for ($i = 0; $i -lt $lines.Count - 1; $i++) {
    if ($lines[$i] -match '^\|\s*Источник\s*\|' -and $lines[$i + 1] -match '^\|[-|]+') {
      for ($j = $i + 2; $j -lt $lines.Count -and $lines[$j].Trim().StartsWith('|'); $j++) {
        $cells = Split-MarkdownRow $lines[$j]
        if ($cells.Count -ge 4 -and $cells[0]) {
          $source = $cells[0].Trim()
          $map[$source] = [pscustomobject]@{
            Source = $source
            Status = $cells[1].Trim()
            What = $cells[2].Trim()
            Used = $cells[3].Trim()
          }
        }
      }
      break
    }
  }
  return $map
}

function New-NsManager {
  param([xml]$Xml)
  $ns = [Xml.XmlNamespaceManager]::new($Xml.NameTable)
  $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
  return $ns
}

function Get-CellText {
  param($Cell)
  if (-not $Cell) { return '' }
  return $Cell.InnerText.Trim()
}

function Get-CellByColumn {
  param($Row, $Ns, [string]$Column)
  return $Row.SelectSingleNode("*[local-name()='c' and starts-with(@r,'$Column')]")
}

function Set-InlineCellText {
  param($Row, $Ns, [string]$Column, [string]$Value)
  $cell = Get-CellByColumn $Row $Ns $Column
  if (-not $cell) { return }
  $t = $cell.SelectSingleNode(".//*[local-name()='t']")
  if (-not $t) { return }
  $t.InnerText = $Value
}

function Canonicalize-SheetNames {
  param([xml]$Xml, $Ns)
  $nameMap = @{
    'System.Xml.XmlElement' = 'hh.ru'
    'исследования и аналитика' = 'Экоспи'
    'Международный валютный фонд (МВФ)' = 'Международный валютный фонд'
    'Про HR и не только | Наталья Володина' = 'Про HR и не только / Наталья Володина'
    'Пульс | PRO. Людей' = 'Пульс / PRO. Людей'
    'ТеДо | ПроНалоги. TaxPro' = 'ТеДо / ПроНалоги. TaxPro'
  }
  foreach ($row in $Xml.DocumentElement.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row'][position()>1]")) {
    $cell = Get-CellByColumn $row $Ns 'B'
    $current = Get-CellText $cell
    if ($nameMap.ContainsKey($current)) {
      Set-InlineCellText $row $Ns 'B' $nameMap[$current]
    }
  }
}

function Set-RowNumber {
  param($Row, $Ns, [int]$Number)
  $Row.SetAttribute('r', [string]$Number)
  foreach ($cell in $Row.SelectNodes("*[local-name()='c']")) {
    $column = $cell.GetAttribute('r') -replace '\d', ''
    $cell.SetAttribute('r', "$column$Number")
  }
}

$aprilMap = Get-AprilSourceMap
$storagePath = (Resolve-Path '.\Хранилище.xlsx').Path
$zip = [IO.Compression.ZipFile]::Open($storagePath, [IO.Compression.ZipArchiveMode]::Update)
try {
  foreach ($sheetName in @('xl/worksheets/sheet4.xml', 'xl/worksheets/sheet5.xml', 'xl/worksheets/sheet7.xml')) {
    [xml]$sheetXml = Read-EntryText $zip $sheetName
    $ns = New-NsManager $sheetXml
    Canonicalize-SheetNames $sheetXml $ns

    if ($sheetName -eq 'xl/worksheets/sheet4.xml') {
      $hasHeyHr = $false
      foreach ($row in $sheetXml.DocumentElement.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row'][position()>1]")) {
        $source = Get-CellText (Get-CellByColumn $row $ns 'B')
        if ($source -eq 'Эй, HR! Антон Платонов') { $hasHeyHr = $true }
        if ($aprilMap.ContainsKey($source)) {
          $item = $aprilMap[$source]
          Set-InlineCellText $row $ns 'E' $item.What
          Set-InlineCellText $row $ns 'N' $item.Used
          Set-InlineCellText $row $ns 'O' $(if ($item.Used -eq 'да') { $item.What } else { '' })
          Set-InlineCellText $row $ns 'P' $item.Status
          Set-InlineCellText $row $ns 'Q' 'восстановлено из готового обзора'
          Set-InlineCellText $row $ns 'S' 'Восстановлено из data\2026-04\monthly_digest.md; презентации: presentations\2026-04_monthly_director_deck.html / presentations\2026-04_monthly_director_deck.pdf'
        }
      }

      if (-not $hasHeyHr) {
        $template = $sheetXml.SelectSingleNode("//*[local-name()='row' and @r='50']")
        $clone = $template.CloneNode($true)
        $rows = @($sheetXml.DocumentElement.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']"))
        $nextRow = ([int]($rows | ForEach-Object { $_.GetAttribute('r') } | Measure-Object -Maximum).Maximum) + 1
        Set-RowNumber $clone $ns $nextRow
        Set-InlineCellText $clone $ns 'B' 'Эй, HR! Антон Платонов'
        Set-InlineCellText $clone $ns 'R' 'https://t.me/Hey_HR'
        $item = $aprilMap['Эй, HR! Антон Платонов']
        Set-InlineCellText $clone $ns 'E' $item.What
        Set-InlineCellText $clone $ns 'N' $item.Used
        Set-InlineCellText $clone $ns 'O' ''
        Set-InlineCellText $clone $ns 'P' $item.Status
        Set-InlineCellText $clone $ns 'Q' 'восстановлено из готового обзора'
        Set-InlineCellText $clone $ns 'S' 'Восстановлено из data\2026-04\monthly_digest.md; презентации: presentations\2026-04_monthly_director_deck.html / presentations\2026-04_monthly_director_deck.pdf'
        $sheetXml.worksheet.sheetData.AppendChild($clone) | Out-Null
        $sheetXml.worksheet.dimension.ref = "A1:AB$nextRow"
      }
    }

    Write-EntryText $zip $sheetName $sheetXml.OuterXml
  }

  [xml]$workbookXml = Read-EntryText $zip 'xl/workbook.xml'
  $workbookText = $workbookXml.OuterXml.Replace('\Дайджест\', '\Обзор РТ\')
  Write-EntryText $zip 'xl/workbook.xml' $workbookText
} finally {
  $zip.Dispose()
}

$companiesPath = (Resolve-Path '.\Компании стратегического наблюдения.xlsx').Path
$companiesZip = [IO.Compression.ZipFile]::Open($companiesPath, [IO.Compression.ZipArchiveMode]::Update)
try {
  $workbookText = Read-EntryText $companiesZip 'xl/workbook.xml'
  $workbookText = $workbookText.Replace('\Дайджест\', '\Обзор РТ\')
  Write-EntryText $companiesZip 'xl/workbook.xml' $workbookText
} finally {
  $companiesZip.Dispose()
}

Write-Output 'STORAGE_AND_XLSX_METADATA_REPAIRED=True'
