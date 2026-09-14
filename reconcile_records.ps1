param(
  [string]$TargetPath = "",
  [string]$ReferencePath = "",
  [string]$OutputPath = "outputs\reconciled_gab_records.xlsx"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

function Read-ZipText($zip, [string]$entryName) {
  $entry = $zip.GetEntry($entryName)
  if ($null -eq $entry) { return $null }
  $stream = $entry.Open()
  try {
    $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)
    return $reader.ReadToEnd()
  }
  finally { $stream.Dispose() }
}

function Get-ColumnIndex([string]$cellRef) {
  $letters = ([regex]::Match($cellRef, '^[A-Z]+')).Value
  $n = 0
  foreach ($ch in $letters.ToCharArray()) {
    $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $n
}

function Get-ColumnName([int]$index) {
  $name = ""
  while ($index -gt 0) {
    $mod = ($index - 1) % 26
    $name = [char]([int][char]'A' + $mod) + $name
    $index = [math]::Floor(($index - 1) / 26)
  }
  return $name
}

function Get-SharedStrings($zip) {
  $list = New-Object System.Collections.Generic.List[string]
  $xmlText = Read-ZipText $zip "xl/sharedStrings.xml"
  if ([string]::IsNullOrWhiteSpace($xmlText)) { return $list }
  [xml]$sst = $xmlText
  foreach ($si in $sst.sst.si) {
    $list.Add([System.Web.HttpUtility]::HtmlDecode($si.InnerText))
  }
  return $list
}

function Get-CellText($cell, $sharedStrings) {
  if ($null -eq $cell) { return "" }
  $t = [string]$cell.t
  if ($t -eq "s") {
    $idx = [int]$cell.v
    if ($idx -ge 0 -and $idx -lt $sharedStrings.Count) { return $sharedStrings[$idx] }
    return ""
  }
  if ($t -eq "inlineStr") { return [string]$cell.is.InnerText }
  if ($null -ne $cell.v) { return [string]$cell.v }
  return [string]$cell.InnerText
}

function Get-SheetMap($zip) {
  [xml]$workbookXml = Read-ZipText $zip "xl/workbook.xml"
  [xml]$relsXml = Read-ZipText $zip "xl/_rels/workbook.xml.rels"
  $relMap = @{}
  foreach ($rel in $relsXml.Relationships.Relationship) {
    $relMap[[string]$rel.Id] = [string]$rel.Target
  }
  $sheets = @{}
  foreach ($sheet in $workbookXml.workbook.sheets.sheet) {
    $rid = $sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $target = $relMap[$rid]
    if ($target -notlike "xl/*") { $target = "xl/$target" }
    $sheets[[string]$sheet.name] = [pscustomobject]@{
      Name = [string]$sheet.name
      Id = [string]$sheet.sheetId
      Path = $target
    }
  }
  return $sheets
}

function Read-SheetRows($zip, $sheetInfo, $sharedStrings) {
  [xml]$xml = Read-ZipText $zip $sheetInfo.Path
  $rows = @()
  foreach ($row in $xml.worksheet.sheetData.row) {
    $values = @{}
    foreach ($cell in $row.c) {
      $values[(Get-ColumnIndex $cell.r)] = Get-CellText $cell $sharedStrings
    }
    $rows += [pscustomobject]@{ RowNumber = [int]$row.r; Values = $values }
  }
  return $rows
}

function Normalize-Text([string]$text) {
  if ([string]::IsNullOrWhiteSpace($text)) { return "" }
  $s = $text.ToLowerInvariant()
  $s = $s -replace 'https?://(www\.)?', ''
  $s = $s -replace '[''`"]', ' '
  $s = $s.Normalize([Text.NormalizationForm]::FormD)
  $chars = New-Object System.Text.StringBuilder
  foreach ($ch in $s.ToCharArray()) {
    $cat = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
    if ($cat -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$chars.Append($ch) }
  }
  $s = $chars.ToString().Normalize([Text.NormalizationForm]::FormC)
  $s = $s -replace ([string][char]0x0111), 'd'
  $s = $s -replace '\b(ky|ki)\s*luc\s*gia\b', ' '
  $s = $s -replace '\b(klg|gs|ts|pgs|ths|nsut|nsnd|ong|ba|anh|chi|thay|hoa si|nha bao|nghe si|giao su|vien si)\b', ' '
  $s = $s -replace '[^a-z0-9]+', ' '
  return (($s -split '\s+' | Where-Object { $_.Length -gt 1 }) -join ' ').Trim()
}

function Get-Tokens([string]$text) {
  $n = Normalize-Text $text
  if ([string]::IsNullOrWhiteSpace($n)) { return @() }
  return @($n -split '\s+' | Where-Object { $_.Length -gt 1 } | Select-Object -Unique)
}

function Token-Score([string]$a, [string]$b) {
  $ta = @(Get-Tokens $a)
  $tb = @(Get-Tokens $b)
  return Token-ScoreFromTokens $ta $tb
}

function Token-ScoreFromTokens($ta, $tb) {
  if ($ta.Count -eq 0 -or $tb.Count -eq 0) { return 0.0 }
  $setB = @{}
  foreach ($t in $tb) { $setB[$t] = $true }
  $common = 0
  foreach ($t in $ta) { if ($setB.ContainsKey($t)) { $common++ } }
  $coverageA = $common / [double]$ta.Count
  $coverageB = $common / [double]$tb.Count
  return [math]::Round(([math]::Max($coverageA, $coverageB) * 0.65 + [math]::Min($coverageA, $coverageB) * 0.35), 4)
}

function Clean-Url([string]$url) {
  if ([string]::IsNullOrWhiteSpace($url)) { return "" }
  return ($url.Trim().ToLowerInvariant() -replace '^http://', 'https://' -replace '/+$','')
}

function New-RefRecord($sheet, $row, $owner, $title, $date, $link) {
  return [pscustomobject]@{
    Sheet = $sheet
    Row = $row
    Owner = [string]$owner
    Title = [string]$title
    Date = [string]$date
    Link = [string]$link
    CleanLink = Clean-Url $link
    OwnerTokens = @(Get-Tokens $owner)
    TitleTokens = @(Get-Tokens $title)
  }
}

function Read-References($path) {
  $zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $path).Path)
  try {
    $shared = Get-SharedStrings $zip
    $sheets = Get-SheetMap $zip
    $refs = @()
    foreach ($name in @("CSDL KLVN", "CSDL KLVW", "CSDL KL CA", "CSDL KLTG")) {
      $rows = Read-SheetRows $zip $sheets[$name] $shared
      foreach ($r in $rows) {
        if ($r.RowNumber -lt 3) { continue }
        if ($name -eq "CSDL KLVN") {
          $refs += New-RefRecord $name $r.RowNumber $r.Values[6] $r.Values[5] $r.Values[4] ""
        } elseif ($name -eq "CSDL KLVW") {
          $refs += New-RefRecord $name $r.RowNumber $r.Values[7] $r.Values[6] $r.Values[5] ""
        } else {
          $refs += New-RefRecord $name $r.RowNumber $r.Values[2] $r.Values[3] $r.Values[5] $r.Values[6]
        }
      }
    }
    return @($refs | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Owner + $_.Title + $_.Link) })
  }
  finally { $zip.Dispose() }
}

function Find-BestMatch($record, $refs) {
  $url = Clean-Url $record.Url
  $recordNameTokens = @(Get-Tokens $record.Name)
  $recordTitleTokens = @(Get-Tokens $record.Title)
  if ($url) {
    $urlMatch = @($refs | Where-Object { $_.CleanLink -and $_.CleanLink -eq $url } | Select-Object -First 1)
    if ($urlMatch.Count -gt 0) {
      return [pscustomobject]@{ Ref = $urlMatch[0]; TitleScore = 1.0; NameScore = (Token-ScoreFromTokens $recordNameTokens $urlMatch[0].OwnerTokens); UrlExact = $true }
    }
  }
  $needle = @{}
  foreach ($t in @($recordNameTokens + $recordTitleTokens)) {
    if ($t.Length -ge 4) { $needle[$t] = $true }
  }
  $candidates = @()
  foreach ($ref in $refs) {
    $hasCommon = $false
    foreach ($t in @($ref.OwnerTokens + $ref.TitleTokens)) {
      if ($needle.ContainsKey($t)) { $hasCommon = $true; break }
    }
    if ($hasCommon) { $candidates += $ref }
  }
  if ($candidates.Count -eq 0) { $candidates = $refs }
  $best = $null
  $bestScore = -1.0
  foreach ($ref in $candidates) {
    $titleScore = Token-ScoreFromTokens $recordTitleTokens $ref.TitleTokens
    $nameScore = Token-ScoreFromTokens $recordNameTokens $ref.OwnerTokens
    $score = ($titleScore * 0.72) + ($nameScore * 0.28)
    if ($score -gt $bestScore) {
      $bestScore = $score
      $best = [pscustomobject]@{ Ref = $ref; TitleScore = $titleScore; NameScore = $nameScore; UrlExact = $false }
    }
  }
  return $best
}

function Build-Note($record, $match) {
  if ($null -eq $match -or $null -eq $match.Ref) {
    return "MISSING: not found in the 4 CSDL tabs."
  }
  $ref = $match.Ref
  $isStrong = $match.UrlExact -or $match.TitleScore -ge 0.82 -or ($match.NameScore -ge 0.78 -and $match.TitleScore -ge 0.55)
  if (-not $isStrong) {
    return ("MISSING/UNCERTAIN: no confident match in the 4 CSDL tabs. Closest: {0}!R{1}; CSDL owner: {2}; CSDL record: {3}; name score {4:P0}, title score {5:P0}." -f $ref.Sheet, $ref.Row, $ref.Owner, $ref.Title, $match.NameScore, $match.TitleScore)
  }
  $issues = @()
  if (-not $match.UrlExact -and $record.Url -and $ref.Link) {
    $issues += ("different link: Data={0}; CSDL={1}" -f $record.Url, $ref.Link)
  } elseif ($record.Url -and -not $ref.Link) {
    $issues += "CSDL has no link for URL check"
  } elseif (-not $record.Url -and $ref.Link) {
    $issues += ("missing link in Data; CSDL has: {0}" -f $ref.Link)
  }
  if ($match.NameScore -lt 0.78) {
    $issues += ("different owner/name: Data='{0}'; CSDL='{1}'" -f $record.Name, $ref.Owner)
  }
  if ($match.TitleScore -lt 0.82) {
    $issues += ("different/missing record title: CSDL='{0}'" -f $ref.Title)
  }
  if ($issues.Count -eq 0) {
    return ("MATCHED with {0}!R{1}." -f $ref.Sheet, $ref.Row)
  }
  return ("REVIEW with {0}!R{1}: {2}" -f $ref.Sheet, $ref.Row, ($issues -join "; "))
}

function Set-Or-Add-Cell($xml, $rowNode, [string]$cellRef, [string]$text) {
  $ns = $xml.DocumentElement.NamespaceURI
  $cell = $null
  foreach ($c in $rowNode.c) {
    if ([string]$c.r -eq $cellRef) { $cell = $c; break }
  }
  if ($null -eq $cell) {
    $cell = $xml.CreateElement("c", $ns)
    [void]$cell.SetAttribute("r", $cellRef)
    $targetCol = Get-ColumnIndex $cellRef
    $inserted = $false
    foreach ($c in @($rowNode.c)) {
      if ((Get-ColumnIndex $c.r) -gt $targetCol) {
        [void]$rowNode.InsertBefore($cell, $c)
        $inserted = $true
        break
      }
    }
    if (-not $inserted) { [void]$rowNode.AppendChild($cell) }
  } else {
    $cell.RemoveAll()
    [void]$cell.SetAttribute("r", $cellRef)
  }
  [void]$cell.SetAttribute("t", "inlineStr")
  $is = $xml.CreateElement("is", $ns)
  $t = $xml.CreateElement("t", $ns)
  [void]$t.SetAttribute("space", "http://www.w3.org/XML/1998/namespace", "preserve")
  $t.InnerText = $text
  [void]$is.AppendChild($t)
  [void]$cell.AppendChild($is)
}

if ([string]::IsNullOrWhiteSpace($TargetPath)) {
  $TargetPath = (Get-ChildItem -LiteralPath "Dulieutruyxuat" -Filter "*.xlsx" | Where-Object { $_.Name -like "*15_24*" } | Select-Object -First 1).FullName
}
if ([string]::IsNullOrWhiteSpace($ReferencePath)) {
  $ReferencePath = (Get-ChildItem -LiteralPath "Dulieutruyxuat" -Filter "*.xlsx" | Where-Object { $_.Name -like "*GAB.xlsx" -and $_.Name -notlike "*15_24*" } | Select-Object -First 1).FullName
}
if (-not $TargetPath -or -not $ReferencePath) { throw "Input workbook not found." }

$refs = Read-References $ReferencePath
$targetFull = (Resolve-Path -LiteralPath $TargetPath).Path
$outputFull = Join-Path (Get-Location) $OutputPath
$outputDir = Split-Path -Parent $outputFull
if (-not (Test-Path -LiteralPath $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
Copy-Item -LiteralPath $targetFull -Destination $outputFull -Force

$workDir = Join-Path $outputDir ("xlsx_work_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workDir | Out-Null
[System.IO.Compression.ZipFile]::ExtractToDirectory($outputFull, $workDir)

$targetZip = [System.IO.Compression.ZipFile]::OpenRead($targetFull)
try {
  $targetShared = Get-SharedStrings $targetZip
  $targetSheets = Get-SheetMap $targetZip
  $dataSheet = @($targetSheets.Values | Where-Object { $_.Name -like "Data_record_01-01-2022_01-01-20*" } | Select-Object -First 1)[0]
  $rows = Read-SheetRows $targetZip $dataSheet $targetShared
}
finally { $targetZip.Dispose() }

$sheetXmlPath = Join-Path $workDir $dataSheet.Path
[xml]$sheetXml = Get-Content -LiteralPath $sheetXmlPath -Encoding UTF8 -Raw
$rowByNumber = @{}
foreach ($rowNode in $sheetXml.worksheet.sheetData.row) { $rowByNumber[[int]$rowNode.r] = $rowNode }

Set-Or-Add-Cell $sheetXml $rowByNumber[1] "H1" "Note"

$summary = [ordered]@{ Matched = 0; Review = 0; Missing = 0; Total = 0 }
foreach ($r in $rows) {
  if ($r.RowNumber -le 1) { continue }
  $record = [pscustomobject]@{
    Row = $r.RowNumber
    RecordId = $r.Values[1]
    GabId = $r.Values[2]
    Name = $r.Values[3]
    Date = $r.Values[4]
    Title = $r.Values[5]
    Url = $r.Values[6]
    Description = $r.Values[7]
  }
  if ([string]::IsNullOrWhiteSpace($record.Name + $record.Title + $record.Url)) { continue }
  $match = Find-BestMatch $record $refs
  $note = Build-Note $record $match
  if ($note.StartsWith("MATCHED")) { $summary.Matched++ }
  elseif ($note.StartsWith("MISSING")) { $summary.Missing++ }
  else { $summary.Review++ }
  $summary.Total++
  Set-Or-Add-Cell $sheetXml $rowByNumber[$r.RowNumber] ("H{0}" -f $r.RowNumber) $note
}

if ($sheetXml.worksheet.dimension) {
  $maxRow = ($rows | Measure-Object -Property RowNumber -Maximum).Maximum
  $sheetXml.worksheet.dimension.ref = "A1:H$maxRow"
}

$settings = [System.Xml.XmlWriterSettings]::new()
$settings.Encoding = [System.Text.UTF8Encoding]::new($false)
$settings.Indent = $false
$writer = [System.Xml.XmlWriter]::Create($sheetXmlPath, $settings)
try { $sheetXml.Save($writer) } finally { $writer.Close() }

$oldOutput = $outputFull
$tmpZip = "$outputFull.tmp"
if (Test-Path -LiteralPath $tmpZip) { Remove-Item -LiteralPath $tmpZip -Force }
if (Test-Path -LiteralPath $oldOutput) { Remove-Item -LiteralPath $oldOutput -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($workDir, $tmpZip)
Move-Item -LiteralPath $tmpZip -Destination $outputFull

"Output`t$outputFull"
"Total`t$($summary.Total)"
"Matched`t$($summary.Matched)"
"Review`t$($summary.Review)"
"Missing`t$($summary.Missing)"
