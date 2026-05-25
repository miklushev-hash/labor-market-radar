Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ErrorActionPreference = 'Stop'
$Root = Get-Location
$StoragePath = Join-Path $Root 'Хранилище.xlsx'
$EventsPath = Join-Path $Root 'События стратегического наблюдения.xlsx'

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

function Write-ZipText($Zip, [string]$Name, [string]$Text) {
  $Entry = Get-ZipEntryByName $Zip $Name
  if ($Entry) { $Entry.Delete() }
  $NewEntry = $Zip.CreateEntry(($Name -replace '/', '\'), [System.IO.Compression.CompressionLevel]::Optimal)
  $Writer = [IO.StreamWriter]::new($NewEntry.Open(), [Text.Encoding]::UTF8)
  try { $Writer.Write($Text) } finally { $Writer.Close() }
}

function Remove-ZipEntry($Zip, [string]$Name) {
  $Entry = Get-ZipEntryByName $Zip $Name
  if ($Entry) { $Entry.Delete() }
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

function XmlEscape([string]$Text) {
  if ($null -eq $Text) { return '' }
  return [Security.SecurityElement]::Escape($Text)
}

function Read-SharedStrings($Zip) {
  $Entry = Get-ZipEntryByName $Zip 'xl/sharedStrings.xml'
  $Shared = @()
  if (-not $Entry) { return $Shared }
  [xml]$Sst = Read-ZipText $Zip 'xl/sharedStrings.xml'
  foreach ($Si in $Sst.sst.si) {
    if ($Si.t) { $Shared += [string]$Si.t }
    else { $Shared += (($Si.r | ForEach-Object { [string]$_.t }) -join '') }
  }
  return $Shared
}

function Get-WorkbookParts($Zip) {
  [xml]$Wb = Read-ZipText $Zip 'xl/workbook.xml'
  [xml]$Rels = Read-ZipText $Zip 'xl/_rels/workbook.xml.rels'
  $RelMap = @{}
  foreach ($R in $Rels.Relationships.Relationship) { $RelMap[$R.Id] = $R.Target }
  return [PSCustomObject]@{ Workbook = $Wb; Rels = $Rels; RelMap = $RelMap }
}

function Get-SheetTargetByName($Zip, [string]$SheetName) {
  $Parts = Get-WorkbookParts $Zip
  foreach ($Sheet in $Parts.Workbook.workbook.sheets.sheet) {
    if ([string]$Sheet.name -eq $SheetName) {
      $Rid = $Sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
      return 'xl/' + $Parts.RelMap[$Rid]
    }
  }
  throw "Sheet not found: $SheetName"
}

function Read-SheetRows($Zip, [string]$SheetName) {
  $Shared = Read-SharedStrings $Zip
  $Target = Get-SheetTargetByName $Zip $SheetName
  [xml]$SheetXml = Read-ZipText $Zip $Target
  $Rows = [System.Collections.ArrayList]::new()
  foreach ($Row in $SheetXml.worksheet.sheetData.row) {
    $Vals = @{}
    foreach ($Cell in $Row.c) {
      $Vals[(Get-ColNum $Cell.r)] = Get-CellText $Cell $Shared
    }
    $Max = if ($Vals.Keys.Count -gt 0) { ($Vals.Keys | Measure-Object -Maximum).Maximum } else { 0 }
    $Out = @()
    for ($I = 1; $I -le $Max; $I++) { $Out += [string]$Vals[$I] }
    [void]$Rows.Add([object[]]$Out)
  }
  return $Rows
}

function New-SheetXml($Rows) {
  $RowCount = [Math]::Max(1, $Rows.Count)
  $MaxCols = 1
  foreach ($R in $Rows) { if ($R.Count -gt $MaxCols) { $MaxCols = $R.Count } }
  $LastCol = Get-ColLetters $MaxCols
  $Sb = [Text.StringBuilder]::new()
  [void]$Sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
  [void]$Sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
  [void]$Sb.AppendLine("  <dimension ref=`"A1:$LastCol$RowCount`"/>")
  [void]$Sb.AppendLine('  <sheetViews><sheetView workbookViewId="0"/></sheetViews>')
  [void]$Sb.AppendLine('  <sheetFormatPr defaultRowHeight="15"/>')
  [void]$Sb.AppendLine('  <sheetData>')
  for ($R = 1; $R -le $Rows.Count; $R++) {
    [void]$Sb.AppendLine("    <row r=`"$R`">")
    $Row = $Rows[$R - 1]
    for ($C = 1; $C -le $MaxCols; $C++) {
      $Ref = (Get-ColLetters $C) + $R
      $Val = if ($C -le $Row.Count) { [string]$Row[$C - 1] } else { '' }
      [void]$Sb.AppendLine("      <c r=`"$Ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(XmlEscape $Val)</t></is></c>")
    }
    [void]$Sb.AppendLine('    </row>')
  }
  [void]$Sb.AppendLine('  </sheetData>')
  [void]$Sb.AppendLine("  <autoFilter ref=`"A1:$LastCol$RowCount`"/>")
  [void]$Sb.AppendLine('</worksheet>')
  return $Sb.ToString()
}

function Add-Or-Replace-Worksheet($WorkbookPath, [string]$SheetName, $Rows) {
  $Zip = [System.IO.Compression.ZipFile]::Open($WorkbookPath, [System.IO.Compression.ZipArchiveMode]::Update)
  try {
    $Parts = Get-WorkbookParts $Zip
    $Wb = $Parts.Workbook
    $Rels = $Parts.Rels
    $RelMap = $Parts.RelMap

    $ExistingSheet = $null
    foreach ($Sheet in $Wb.workbook.sheets.sheet) {
      if ([string]$Sheet.name -eq $SheetName) { $ExistingSheet = $Sheet; break }
    }
    if ($ExistingSheet) {
      $Rid = $ExistingSheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
      $Target = 'xl/' + $RelMap[$Rid]
      Write-ZipText $Zip $Target (New-SheetXml $Rows)
      return
    }

    $SheetNums = @()
    foreach ($Entry in $Zip.Entries) {
      if ($Entry.FullName -match 'xl[\\/]worksheets[\\/]sheet(\d+)\.xml') { $SheetNums += [int]$Matches[1] }
    }
    $NewSheetNum = (($SheetNums | Measure-Object -Maximum).Maximum + 1)
    $NewTarget = "worksheets/sheet$NewSheetNum.xml"
    $NewFullName = "xl/worksheets/sheet$NewSheetNum.xml"

    $ExistingIds = @()
    foreach ($Sheet in $Wb.workbook.sheets.sheet) { $ExistingIds += [int]$Sheet.sheetId }
    $NewSheetId = (($ExistingIds | Measure-Object -Maximum).Maximum + 1)

    $ExistingRidNums = @()
    foreach ($Rel in $Rels.Relationships.Relationship) {
      if ($Rel.Id -match '^rId(\d+)$') { $ExistingRidNums += [int]$Matches[1] }
    }
    $NewRid = 'rId' + (($ExistingRidNums | Measure-Object -Maximum).Maximum + 1)

    $NewSheet = $Wb.CreateElement('sheet', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $NewSheet.SetAttribute('name', $SheetName)
    $NewSheet.SetAttribute('sheetId', [string]$NewSheetId)
    $NewSheet.SetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', $NewRid)
    [void]$Wb.workbook.sheets.AppendChild($NewSheet)

    $NewRel = $Rels.CreateElement('Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $NewRel.SetAttribute('Id', $NewRid)
    $NewRel.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet')
    $NewRel.SetAttribute('Target', $NewTarget)
    [void]$Rels.Relationships.AppendChild($NewRel)

    [xml]$Types = Read-ZipText $Zip '[Content_Types].xml'
    $OverrideExists = $false
    foreach ($Override in $Types.Types.Override) {
      if ($Override.PartName -eq "/xl/worksheets/sheet$NewSheetNum.xml") { $OverrideExists = $true }
    }
    if (-not $OverrideExists) {
      $NewOverride = $Types.CreateElement('Override', 'http://schemas.openxmlformats.org/package/2006/content-types')
      $NewOverride.SetAttribute('PartName', "/xl/worksheets/sheet$NewSheetNum.xml")
      $NewOverride.SetAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
      [void]$Types.Types.AppendChild($NewOverride)
    }

    Write-ZipText $Zip $NewFullName (New-SheetXml $Rows)
    Write-ZipText $Zip 'xl/workbook.xml' $Wb.OuterXml
    Write-ZipText $Zip 'xl/_rels/workbook.xml.rels' $Rels.OuterXml
    Write-ZipText $Zip '[Content_Types].xml' $Types.OuterXml
  } finally {
    $Zip.Dispose()
  }
}

function Remove-Worksheet($WorkbookPath, [string]$SheetName) {
  $Zip = [System.IO.Compression.ZipFile]::Open($WorkbookPath, [System.IO.Compression.ZipArchiveMode]::Update)
  try {
    $Parts = Get-WorkbookParts $Zip
    $Wb = $Parts.Workbook
    $Rels = $Parts.Rels
    $RelMap = $Parts.RelMap

    $TargetSheet = $null
    foreach ($Sheet in $Wb.workbook.sheets.sheet) {
      if ([string]$Sheet.name -eq $SheetName) { $TargetSheet = $Sheet; break }
    }
    if (-not $TargetSheet) { return }

    $Rid = $TargetSheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $Target = $RelMap[$Rid]
    $FullTarget = 'xl/' + $Target
    $PartName = '/' + $FullTarget

    [void]$Wb.workbook.sheets.RemoveChild($TargetSheet)
    foreach ($Rel in @($Rels.Relationships.Relationship)) {
      if ($Rel.Id -eq $Rid) { [void]$Rels.Relationships.RemoveChild($Rel) }
    }

    [xml]$Types = Read-ZipText $Zip '[Content_Types].xml'
    foreach ($Override in @($Types.Types.Override)) {
      if ($Override.PartName -eq $PartName) { [void]$Types.Types.RemoveChild($Override) }
    }

    Remove-ZipEntry $Zip $FullTarget
    Write-ZipText $Zip 'xl/workbook.xml' $Wb.OuterXml
    Write-ZipText $Zip 'xl/_rels/workbook.xml.rels' $Rels.OuterXml
    Write-ZipText $Zip '[Content_Types].xml' $Types.OuterXml
  } finally {
    $Zip.Dispose()
  }
}

$EventsZip = [System.IO.Compression.ZipFile]::OpenRead($EventsPath)
try {
  $JournalRows = Read-SheetRows $EventsZip 'Журнал_событий'
} finally {
  $EventsZip.Dispose()
}

Add-Or-Replace-Worksheet $StoragePath 'Журнал_событий' $JournalRows
Remove-Worksheet $EventsPath 'Журнал_событий'

[PSCustomObject]@{
  MovedRows = $JournalRows.Count
  Storage = $StoragePath
  Events = $EventsPath
} | Format-List
