# ============================================================
# Generate RPT_Import_Map.xlsx from folder structure under -RptRoot.
# Folder name pattern: <8-char DocCode>__<DocName>
#   TypeCode = first 4 chars of folder name
#   DocCode  = first 8 chars of folder name (TypeCode + 4-digit seq)
#   LayoutName = portion after '__'
# Writes sheet 'RPT_MAP' with columns: No, DocCode, TypeCode,
# RPT_FileName, RPT_Folder, LayoutName, Note
# ============================================================
param(
    [string]$RptRoot = "C:\GitHub\Enconfund\FORM",
    [string]$OutFile = "$PSScriptRoot\..\Config\RPT_Import_Map.xlsx",
    [string[]]$ExcludeFolders = @("sap-b1-layout-import-SQL-")
)

$ErrorActionPreference = "Stop"

Write-Host "=== Generate-MapExcel ===" -ForegroundColor Cyan
Write-Host "RptRoot : $RptRoot"
Write-Host "OutFile : $OutFile"
Write-Host ""

if (-not (Test-Path $RptRoot)) { throw "RptRoot not found: $RptRoot" }

# ---------- 1. Scan folders ----------
$folders = Get-ChildItem $RptRoot -Directory |
           Where-Object { $ExcludeFolders -notcontains $_.Name } |
           Sort-Object Name

$rows = @()
$idx = 0
foreach ($f in $folders) {
    $name = $f.Name
    if ($name -notmatch '^([A-Z0-9]{4})([0-9]{4})__(.+)$') {
        Write-Host ("  SKIP bad name : {0}" -f $name) -ForegroundColor Yellow
        continue
    }
    $typeCode   = $matches[1]
    $docCode    = $matches[1] + $matches[2]
    $layoutName = $matches[3]
    $rpt = Get-ChildItem $f.FullName -Filter "*.rpt" -File | Select-Object -First 1
    if (-not $rpt) {
        Write-Host ("  SKIP no .rpt  : {0}" -f $name) -ForegroundColor Yellow
        continue
    }
    $idx++
    $rows += [PSCustomObject]@{
        No           = $idx
        DocCode      = $docCode
        TypeCode     = $typeCode
        RPT_FileName = $rpt.Name
        RPT_Folder   = $name
        LayoutName   = $layoutName
        Note         = ""
    }
}

Write-Host ("Built {0} rows" -f $rows.Count) -ForegroundColor Green

# ---------- 2. Group summary ----------
Write-Host "`nTypeCode groups:" -ForegroundColor Cyan
$rows | Group-Object TypeCode | Sort-Object Name |
    ForEach-Object { "  {0,-6} {1,3} layouts" -f $_.Name, $_.Count }

# ---------- 3. Write Excel via OpenXML zip (no Excel COM dependency) ----------
# Collect strings (cells use shared-string indices for text)
$ssList    = New-Object System.Collections.Generic.List[string]
$ssIdx     = @{}
function SS-Add([string]$s) {
    if (-not $ssIdx.ContainsKey($s)) {
        $ssIdx[$s] = $ssList.Count
        $ssList.Add($s)
    }
    return $ssIdx[$s]
}

# Header row
$headers = @("No","DocCode","TypeCode","RPT_FileName","RPT_Folder","LayoutName","Note")
$headerIdx = @()
foreach ($h in $headers) { $headerIdx += SS-Add $h }

# Build sheet1 (RPT_MAP) XML
function XmlEscape([string]$s) {
    if ($null -eq $s) { return "" }
    return [System.Security.SecurityElement]::Escape($s)
}

$cellsXml = New-Object System.Text.StringBuilder
[void]$cellsXml.Append('<row r="1">')
$col = 0
foreach ($h in $headers) {
    $col++
    $ref = ([char](64+$col)) + "1"
    [void]$cellsXml.AppendFormat('<c r="{0}" t="s"><v>{1}</v></c>', $ref, $headerIdx[$col-1])
}
[void]$cellsXml.Append('</row>')

$rowNum = 1
foreach ($r in $rows) {
    $rowNum++
    [void]$cellsXml.AppendFormat('<row r="{0}">', $rowNum)
    $vals = @($r.No, $r.DocCode, $r.TypeCode, $r.RPT_FileName, $r.RPT_Folder, $r.LayoutName, $r.Note)
    for ($i=0; $i -lt $vals.Count; $i++) {
        $ref = ([char](65+$i)) + $rowNum
        $v = $vals[$i]
        if ($i -eq 0) {
            # No -> numeric
            [void]$cellsXml.AppendFormat('<c r="{0}"><v>{1}</v></c>', $ref, $v)
        } else {
            $sIdx = SS-Add ([string]$v)
            [void]$cellsXml.AppendFormat('<c r="{0}" t="s"><v>{1}</v></c>', $ref, $sIdx)
        }
    }
    [void]$cellsXml.Append('</row>')
}

$sheet1Xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:G$rowNum"/>
<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>
<col min="1" max="1" width="5"/>
<col min="2" max="2" width="11"/>
<col min="3" max="3" width="9"/>
<col min="4" max="4" width="60"/>
<col min="5" max="5" width="50"/>
<col min="6" max="6" width="40"/>
<col min="7" max="7" width="30"/>
</cols>
<sheetData>
$($cellsXml.ToString())
</sheetData>
</worksheet>
"@
# (entries written below directly into the zip with forward-slash paths)

# sharedStrings.xml
$ssXml = New-Object System.Text.StringBuilder
[void]$ssXml.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
[void]$ssXml.AppendFormat('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{0}" uniqueCount="{0}">', $ssList.Count)
foreach ($s in $ssList) {
    [void]$ssXml.AppendFormat('<si><t xml:space="preserve">{0}</t></si>', (XmlEscape $s))
}
[void]$ssXml.Append('</sst>')
$sharedStringsXml = $ssXml.ToString()

# workbook.xml
$workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="RPT_MAP" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>
"@

# xl/_rels/workbook.xml.rels
$wbRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
"@

# _rels/.rels
$rootRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"@

# [Content_Types].xml
$ct = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
"@

# Backup existing file
if (Test-Path $OutFile) {
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $bak = "$OutFile.$ts.bak"
    Copy-Item $OutFile $bak -Force
    Write-Host ("Backed up existing -> {0}" -f $bak) -ForegroundColor DarkGray
}

# Build xlsx zip with explicit forward-slash entry names (OPC spec)
$outDir = Split-Path $OutFile -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (Test-Path $OutFile) { Remove-Item $OutFile -Force }

$fs = [System.IO.File]::Open($OutFile, [System.IO.FileMode]::CreateNew)
try {
    $archive = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        function Write-ZipEntry($archive, [string]$name, [string]$content) {
            $entry = $archive.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $w = New-Object System.IO.StreamWriter($entry.Open(), [System.Text.UTF8Encoding]::new($false))
            try { $w.Write($content) } finally { $w.Dispose() }
        }
        Write-ZipEntry $archive '[Content_Types].xml'         $ct
        Write-ZipEntry $archive '_rels/.rels'                  $rootRels
        Write-ZipEntry $archive 'xl/workbook.xml'              $workbookXml
        Write-ZipEntry $archive 'xl/_rels/workbook.xml.rels'   $wbRels
        Write-ZipEntry $archive 'xl/sharedStrings.xml'         $sharedStringsXml
        Write-ZipEntry $archive 'xl/worksheets/sheet1.xml'     $sheet1Xml
    } finally { $archive.Dispose() }
} finally { $fs.Dispose() }

Write-Host ""
Write-Host ("=== DONE: {0} ({1} rows) ===" -f $OutFile, $rows.Count) -ForegroundColor Green
