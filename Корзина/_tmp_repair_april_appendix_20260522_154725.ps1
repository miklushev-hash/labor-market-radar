Add-Type -AssemblyName System.Web

$MarkdownPath = 'data\2026-04\monthly_digest.md'
$HtmlPath = 'presentations\2026-04_monthly_director_deck.html'
$BadNumericSources = @('22', '15', '16', '18', '21')

function Decode-BrokenRow([string]$Line) {
  $Decoded = $Line.Replace('\\|', '').Replace('\\', '')
  $Parts = @($Decoded -split '(?<!\\)\|')
  $Cells = @($Parts | Select-Object -Skip 1 | Select-Object -SkipLast 1 | ForEach-Object { $_.Trim() })

  if ($Cells.Count -eq 5 -and $Cells[0] -eq 'Про HR и не только') {
    $Cells = @('Про HR и не только / Наталья Володина', $Cells[2], $Cells[3], $Cells[4])
  }
  elseif ($Cells.Count -eq 5 -and $Cells[0] -eq 'ТеДо') {
    $Cells = @('ТеДо / ПроНалоги. TaxPro', $Cells[2], $Cells[3], $Cells[4])
  }
  elseif ($Cells.Count -eq 5 -and $Cells[0] -eq 'Пульс') {
    $Cells = @('Пульс / PRO. Людей', $Cells[2], $Cells[3], $Cells[4])
  }

  if ($Cells.Count -ne 4) {
    throw "Unexpected appendix row shape: $Line"
  }

  $Cells[0] = $Cells[0].Replace('Международный волютный фонд', 'Международный валютный фонд')
  $Cells[0] = $Cells[0].Replace('Комерсант', 'Коммерсантъ')
  $Cells[0] = $Cells[0].Replace('Т-эдвайзер', 'TAdviser')
  $Cells[0] = $Cells[0].Replace('Теленд Код', 'Talent Code')
  return ,$Cells
}

function Escape-MarkdownCell([string]$Text) {
  return ($Text -replace '\|', '\|')
}

$Lines = [System.Collections.Generic.List[string]](Get-Content -LiteralPath $MarkdownPath -Encoding UTF8)
$HeaderIndex = $Lines.IndexOf('| Источник | Статус за апрель | Что найдено или почему данных нет | Использован в выпуске |')
if ($HeaderIndex -lt 0) { throw 'April appendix header not found in Markdown.' }
$SeparatorIndex = $HeaderIndex + 1
$RowStart = $SeparatorIndex + 1
$RowEnd = $RowStart
while ($RowEnd -lt $Lines.Count -and $Lines[$RowEnd].StartsWith('| ')) { $RowEnd++ }

$Rows = [System.Collections.ArrayList]::new()
for ($I = $RowStart; $I -lt $RowEnd; $I++) {
  $Cells = Decode-BrokenRow $Lines[$I]
  if ($BadNumericSources -contains $Cells[0]) { continue }
  [void]$Rows.Add([object[]]$Cells)
}

$NewMarkdownRows = @()
foreach ($Cells in $Rows) {
  $NewMarkdownRows += ('| ' + (($Cells | ForEach-Object { Escape-MarkdownCell $_ }) -join ' | ') + ' |')
}
$Lines.RemoveRange($RowStart, $RowEnd - $RowStart)
$Lines.InsertRange($RowStart, [string[]]$NewMarkdownRows)
Set-Content -LiteralPath $MarkdownPath -Value $Lines -Encoding UTF8

$HtmlRows = [System.Text.StringBuilder]::new()
foreach ($Cells in $Rows) {
  [void]$HtmlRows.AppendLine('<tr>')
  foreach ($Cell in $Cells) {
    [void]$HtmlRows.AppendLine('<td>' + [System.Web.HttpUtility]::HtmlEncode($Cell) + '</td>')
  }
  [void]$HtmlRows.AppendLine('</tr>')
}

$Html = Get-Content -LiteralPath $HtmlPath -Encoding UTF8 -Raw
$Replacement = @"
<h2>Приложение по источникам</h2>
<div class="table-wrap"><table><thead><tr>
<th>Источник</th>
<th>Статус за апрель</th>
<th>Что найдено или почему данных нет</th>
<th>Использован в выпуске</th>
</tr></thead><tbody>
$($HtmlRows.ToString().TrimEnd())
</tbody></table></div>
"@
$Pattern = '(?s)<h2>Приложение по источникам</h2>\s*<div class="table-wrap"><table>.*?</tbody></table></div>'
$NewHtml = [regex]::Replace($Html, $Pattern, $Replacement, 1)
if ($NewHtml -eq $Html) { throw 'April appendix table not replaced in HTML.' }
Set-Content -LiteralPath $HtmlPath -Value $NewHtml -Encoding UTF8

[pscustomobject]@{
  RepairedRows = $Rows.Count
  Markdown = $MarkdownPath
  Html = $HtmlPath
} | Format-List
