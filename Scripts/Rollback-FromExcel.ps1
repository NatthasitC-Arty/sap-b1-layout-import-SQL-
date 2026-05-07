# ============================================================
# Rollback: Delete layouts in RDOC that match rows in Excel mapping
# Reads RPT_Import_Map.xlsx -> deletes matching rows by (DocName + TypeCode)
# ============================================================
param(
    [string]$Server     = "SLD-C072",
    [string]$CompanyDB  = "SBO_SDA",
    [string]$DBUser     = "sa",
    [string]$DBPassword = "1q2w3e4r",
    [string]$MapFile    = "C:\SDA\SDA\Form-Layout\ImportLayouts\Config\RPT_Import_Map.xlsx",
    [string]$Author     = "SDA",
    [switch]$UseFileNameAsDocName,
    [switch]$DryRun,
    [switch]$Force
)

# Same ObjectType -> TypeCode map as Import_SQL_Direct.ps1
$TypeCodeMap = @{
    "30"         = "JDT2"
    "23"         = "QUT2"
    "17"         = "RDR2"
    "15"         = "DLN2"
    "16"         = "RDN2"
    "203"        = "DPI2"
    "13"         = "INV2"
    "14"         = "RIN2"
    "1470000113" = "PRQ2"
    "540000405"  = "PQT2"
    "22"         = "POR2"
    "20"         = "PDN2"
    "21"         = "RPD2"
    "204"        = "DPO2"
    "18"         = "PCH2"
    "19"         = "RPC2"
    "69"         = "IPF1"
    "24"         = "RCT1"
    "46"         = "VPM1"
    "59"         = "IGN1"
    "60"         = "IGE1"
    "67"         = "WTR1"
    "1250000001" = "WTQ1"
    "1470000065" = "INC1"
    "162"        = ""
    "202"        = "WOR1"
}

function Read-MapExcel {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    function ColLetter-ToIndex([string]$letters) {
        $n = 0
        foreach ($c in $letters.ToCharArray()) { $n = $n * 26 + ([int][char]$c - 64) }
        return $n
    }

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $shared = @()
        $ssEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" } | Select-Object -First 1
        if ($ssEntry) {
            $sr = New-Object System.IO.StreamReader($ssEntry.Open())
            [xml]$ssXml = $sr.ReadToEnd(); $sr.Close()
            $ns = New-Object System.Xml.XmlNamespaceManager($ssXml.NameTable)
            $ns.AddNamespace("x","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            foreach ($si in $ssXml.SelectNodes("//x:si",$ns)) {
                $txt = ""
                foreach ($t in $si.SelectNodes(".//x:t",$ns)) { $txt += $t.InnerText }
                $shared += ,$txt
            }
        }

        $wbEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/workbook.xml" } | Select-Object -First 1
        $sr = New-Object System.IO.StreamReader($wbEntry.Open())
        [xml]$wbXml = $sr.ReadToEnd(); $sr.Close()
        $nsw = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
        $nsw.AddNamespace("x","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $nsw.AddNamespace("r","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        $sheetNode = $wbXml.SelectSingleNode("//x:sheet[@name='RPT_MAP']",$nsw)
        if (-not $sheetNode) { throw "Sheet 'RPT_MAP' not found in $Path" }
        $rid = $sheetNode.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        $relsEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/_rels/workbook.xml.rels" } | Select-Object -First 1
        $sr = New-Object System.IO.StreamReader($relsEntry.Open())
        [xml]$relsXml = $sr.ReadToEnd(); $sr.Close()
        $target = ($relsXml.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target
        if ($target -notmatch "^/") { $target = "xl/$target" } else { $target = $target.TrimStart('/') }

        $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq $target } | Select-Object -First 1
        $sr = New-Object System.IO.StreamReader($sheetEntry.Open())
        [xml]$shXml = $sr.ReadToEnd(); $sr.Close()
        $nss = New-Object System.Xml.XmlNamespaceManager($shXml.NameTable)
        $nss.AddNamespace("x","http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        $list = @()
        foreach ($row in $shXml.SelectNodes("//x:sheetData/x:row",$nss)) {
            $rowIdx = [int]$row.r
            if ($rowIdx -lt 2) { continue }
            $cells = @{}
            foreach ($c in $row.SelectNodes("x:c",$nss)) {
                $ref = $c.r
                $letters = ($ref -replace '[0-9]','')
                $colIdx = ColLetter-ToIndex $letters
                $t = $c.t
                $vNode = $c.SelectSingleNode("x:v",$nss)
                $isNode = $c.SelectSingleNode("x:is",$nss)
                $val = $null
                if ($t -eq "s" -and $vNode) {
                    $idx = [int]$vNode.InnerText
                    if ($idx -lt $shared.Count) { $val = $shared[$idx] }
                } elseif ($t -eq "inlineStr" -and $isNode) {
                    $val = ""
                    foreach ($tt in $isNode.SelectNodes(".//x:t",$nss)) { $val += $tt.InnerText }
                } elseif ($vNode) {
                    $val = $vNode.InnerText
                }
                $cells[$colIdx] = $val
            }
            $item = [PSCustomObject]@{
                No           = $cells[1]
                RPT_FileName = $cells[3]
                ObjectType   = $cells[8]
                LayoutName   = $cells[10]
            }
            if ($item.RPT_FileName) { $list += $item }
        }
        return $list
    } finally {
        $zip.Dispose()
    }
}

Write-Host "=== Rollback layouts from Excel mapping ===" -ForegroundColor Cyan
Write-Host "MapFile: $MapFile"
Write-Host "Author : $Author"
Write-Host ""

$mapRows = Read-MapExcel -Path $MapFile
Write-Host "Loaded $($mapRows.Count) mapping rows"

$cs = "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10;"
$conn = New-Object System.Data.SqlClient.SqlConnection $cs
$conn.Open()

$toDelete = @()
$skipped = 0
foreach ($row in $mapRows) {
    $objType = [string]$row.ObjectType
    if ($UseFileNameAsDocName) {
        $layoutName = [System.IO.Path]::GetFileNameWithoutExtension($row.RPT_FileName)
    } else {
        $layoutName = if ([string]::IsNullOrWhiteSpace($row.LayoutName)) { $row.RPT_FileName } else { [string]$row.LayoutName }
    }
    if ([string]::IsNullOrWhiteSpace($objType) -or $objType -eq "-") { $skipped++; continue }
    $typeCode = $TypeCodeMap[$objType]
    if ([string]::IsNullOrWhiteSpace($typeCode)) { $skipped++; continue }

    # Match on DocName + TypeCode only (Author ignored, mirrors Import dedup).
    # Capture ALL matches so duplicates that pre-date the dedup fix get cleaned too.
    $chk = $conn.CreateCommand()
    $chk.CommandText = "SELECT DocCode FROM RDOC WHERE DocName=@n AND TypeCode=@t"
    [void]$chk.Parameters.AddWithValue("@n", $layoutName)
    [void]$chk.Parameters.AddWithValue("@t", $typeCode)
    $rdr = $chk.ExecuteReader()
    while ($rdr.Read()) {
        $toDelete += [PSCustomObject]@{ No=$row.No; DocCode=[string]$rdr["DocCode"]; TypeCode=$typeCode; DocName=$layoutName }
    }
    $rdr.Close()
}

Write-Host ""
Write-Host "=== Preview: $($toDelete.Count) row(s) to delete ===" -ForegroundColor Yellow
$toDelete | Format-Table No,DocCode,TypeCode,DocName -AutoSize

if ($toDelete.Count -eq 0) {
    Write-Host "Nothing to delete." -ForegroundColor Green
    $conn.Close()
    return
}

if ($DryRun) {
    Write-Host "DryRun mode - no changes made." -ForegroundColor Cyan
    $conn.Close()
    return
}

if (-not $Force) {
    $ans = Read-Host "Delete these $($toDelete.Count) rows from RDOC? Type 'yes' to confirm"
    if ($ans -ne "yes") {
        Write-Host "Cancelled." -ForegroundColor Yellow
        $conn.Close()
        return
    }
}

$deleted = 0
foreach ($row in $toDelete) {
    $del = $conn.CreateCommand()
    $del.CommandText = "DELETE FROM RDOC WHERE DocCode=@c"
    [void]$del.Parameters.AddWithValue("@c", $row.DocCode)
    $n = $del.ExecuteNonQuery()
    if ($n -gt 0) {
        Write-Host ("DELETED {0} ({1})" -f $row.DocCode, $row.DocName) -ForegroundColor Green
        $deleted++
    }
}
$conn.Close()

Write-Host ""
Write-Host "=== Summary: Deleted=$deleted / Matched=$($toDelete.Count) / Skipped(no mapping)=$skipped ===" -ForegroundColor Cyan
