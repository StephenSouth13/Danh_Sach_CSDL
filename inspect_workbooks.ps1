param(
  [Parameter(Mandatory=$true)][string]$Path,
  [int]$HeaderRows = 5,
  [int]$DataRows = 3
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Get-ColIndex([string]$cellRef) {
  $letters = ([regex]::Match($cellRef, '^[A-Z]+')).Value
  $n = 0
  foreach ($ch in $letters.ToCharArray()) {
    $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $n
}

function Get-CellText($cell, $sharedStrings) {
  if ($null -eq $cell) { return "" }
  $t = [string]$cell.t
  if ($t -eq "s") {
    $idx = [int]$cell.v
    if ($idx -ge 0 -and $idx -lt $sharedStrings.Count) { return $sharedStrings[$idx] }
    return ""
  }
  if ($t -eq "inlineStr") {
    return (($cell.is.t | ForEach-Object { $_.'#text' }) -join "")
  }
  if ($null -ne $cell.v) { return [string]$cell.v }
  return ""
}

function Get-SheetMap($zip) {
  [xml]$workbookXml = [System.IO.StreamReader]::new($zip.GetEntry("xl/workbook.xml").Open()).ReadToEnd()
  [xml]$relsXml = [System.IO.StreamReader]::new($zip.GetEntry("xl/_rels/workbook.xml.rels").Open()).ReadToEnd()
  $relMap = @{}
  foreach ($rel in $relsXml.Relationships.Relationship) {
    $relMap[[string]$rel.Id] = [string]$rel.Target
  }
  $sheets = @()
  foreach ($sheet in $workbookXml.workbook.sheets.sheet) {
    $rid = $sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $target = $relMap[$rid]
    if ($target -notlike "xl/*") { $target = "xl/$target" }
    $sheets += [pscustomobject]@{ Name = [string]$sheet.name; Id = [string]$sheet.sheetId; Path = $target }
  }
  return $sheets
}

function Get-SharedStrings($zip) {
  $entry = $zip.GetEntry("xl/sharedStrings.xml")
  $list = New-Object System.Collections.Generic.List[string]
  if ($null -eq $entry) { return $list }
  [xml]$sst = [System.IO.StreamReader]::new($entry.Open()).ReadToEnd()
  foreach ($si in $sst.sst.si) {
    $texts = @()
    if ($si.t) { $texts += [string]$si.t }
    if ($si.r) {
      foreach ($r in $si.r) {
        if ($r.t) { $texts += [string]$r.t }
      }
    }
    $list.Add(($texts -join ""))
  }
  return $list
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$full = (Resolve-Path -LiteralPath $Path).Path
$zip = [System.IO.Compression.ZipFile]::OpenRead($full)
try {
  $shared = Get-SharedStrings $zip
  $sheets = Get-SheetMap $zip
  "FILE`t$full"
  foreach ($s in $sheets) {
    "SHEET`t$($s.Id)`t$($s.Name)`t$($s.Path)"
    $entry = $zip.GetEntry($s.Path)
    if ($null -eq $entry) { continue }
    [xml]$xml = [System.IO.StreamReader]::new($entry.Open()).ReadToEnd()
    $rows = @($xml.worksheet.sheetData.row | Select-Object -First ($HeaderRows + $DataRows))
    foreach ($row in $rows) {
      $cells = @()
      foreach ($c in $row.c) {
        $cells += ("{0}:{1}" -f $c.r, (Get-CellText $c $shared))
      }
      "ROW`t$($s.Name)`t$($row.r)`t$($cells -join ' | ')"
    }
  }
}
finally {
  $zip.Dispose()
}
