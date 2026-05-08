# ============================================================
# Backup the RDOC table to disk (BCP-style export, HANA edition)
# Outputs: RDOC_Backup_<timestamp>.csv  (index)
#          RDOC_Backup_<timestamp>.bak  (binary blob dump)
# ============================================================
param(
    [string]$Server     = "10.10.10.109:30015",
    [string]$CompanyDB  = "SBO_ENCONFUND_TRAINING",   # HANA schema
    [string]$DBUser     = "SYSTEM",
    [string]$DBPassword = "",
    [string]$OutDir     = "$PSScriptRoot\..\Backups"
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\DB-HANA.ps1"

$schemaQ = Get-DBQuoteIdent $CompanyDB

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$outFile = Join-Path $OutDir "RDOC_Backup_$ts.bak"

$conn = New-DBConnection -Server $Server -Database $CompanyDB -User $DBUser -Password $DBPassword

$cmd = $conn.CreateCommand()
Set-DBCommandText $cmd "SELECT COUNT(*) FROM $schemaQ.RDOC"
$total = Invoke-DBScalar $cmd
Write-Host "RDOC currently has $total rows" -ForegroundColor Cyan

# Lightweight CSV index of every row (DocCode + identifying metadata)
$cmdIdx = $conn.CreateCommand()
Set-DBCommandText $cmdIdx "SELECT DocCode, TypeCode, DocName, Category, LENGTH(Template) AS Bytes, RptHash FROM $schemaQ.RDOC"
$rdr = Invoke-DBReader $cmdIdx
$rows = New-Object System.Collections.Generic.List[object]
while ($rdr.Read()) {
    $rows.Add([PSCustomObject]@{
        DocCode  = [string]$rdr['DocCode']
        TypeCode = [string]$rdr['TypeCode']
        DocName  = [string]$rdr['DocName']
        Category = [string]$rdr['Category']
        Bytes    = [int]$rdr['Bytes']
        RptHash  = [string]$rdr['RptHash']
    })
}
$rdr.Close()

$csvFile = $outFile -replace "\.bak$", ".csv"
$rows | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
Write-Host ("Saved index of {0} rows -> {1}" -f $rows.Count, $csvFile) -ForegroundColor Green

# Binary dump: magic header + per-row {8-byte DocCode | 4-byte length | bytes}
$fs = [System.IO.File]::OpenWrite($outFile)
$bw = New-Object System.IO.BinaryWriter $fs
$bw.Write([byte[]]([byte]0x52,0x44,0x4F,0x43,0x42,0x4B,0x50,0x31)) # magic "RDOCBKP1"
$bw.Write([int32]$rows.Count)

$cmdBlob = $conn.CreateCommand()
Set-DBCommandText $cmdBlob "SELECT DocCode, Template FROM $schemaQ.RDOC WHERE Template IS NOT NULL"
$reader = Invoke-DBReader $cmdBlob
$savedBinary = 0
while ($reader.Read()) {
    $code = [string]$reader['DocCode']
    $blob = $reader['Template']
    if ($blob -is [byte[]]) {
        $codeBytes = [System.Text.Encoding]::ASCII.GetBytes($code.PadRight(8))
        $bw.Write($codeBytes)
        $bw.Write([int32]$blob.Length)
        $bw.Write($blob)
        $savedBinary++
    }
}
$reader.Close()
$bw.Close()
$fs.Close()
$conn.Close()

Write-Host ("Saved {0} binary templates -> {1}" -f $savedBinary, $outFile) -ForegroundColor Green
Write-Host ""
Write-Host "Backup complete:" -ForegroundColor Yellow
Write-Host ("  Index : {0}" -f $csvFile)
Write-Host ("  Binary: {0} ({1} bytes)" -f $outFile, (Get-Item $outFile).Length)
