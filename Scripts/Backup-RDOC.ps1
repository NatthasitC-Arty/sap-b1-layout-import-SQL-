# ============================================================
# Backup RDOC table to disk (BCP-style export)
# Creates: RDOC_Backup_<timestamp>.csv + RDOC_Backup_<timestamp>.bak (binary)
# Use Restore-RDOC-Backup.ps1 to roll back.
# Works on MSSQL and HANA via the DB plugin.
# ============================================================
param(
    [string]$Server     = "SLD-C072",
    [string]$CompanyDB  = "SBO_SDA",
    [string]$DBUser     = "sa",
    [string]$DBPassword = "1q2w3e4r",
    [string]$OutDir     = "$PSScriptRoot\..\Backups",
    [ValidateSet("MSSQL","HANA")]
    [string]$DBEngine   = "MSSQL"
)

# Load DB plugin
. "$PSScriptRoot\DB-$DBEngine.ps1"

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$outFile = Join-Path $OutDir "RDOC_Backup_$ts.bak"

$conn = New-DBConnection -Server $Server -Database $CompanyDB -User $DBUser -Password $DBPassword
$cmd = $conn.CreateCommand()
$cmd.CommandText = Convert-DBSql "SELECT COUNT(*) FROM RDOC"
$total = $cmd.ExecuteScalar()
Write-Host "RDOC currently has $total rows" -ForegroundColor Cyan

# Index of existing rows -> CSV (DocCode, TypeCode, DocName, Category, Bytes, RptHash)
$cmd.CommandText = Convert-DBSql "SELECT DocCode, TypeCode, DocName, Category, DATALENGTH(Template) AS Bytes, RptHash FROM RDOC"
$reader = $cmd.ExecuteReader()
$rows = New-Object System.Collections.ArrayList
while ($reader.Read()) {
    [void]$rows.Add([PSCustomObject]@{
        DocCode  = [string]$reader["DocCode"]
        TypeCode = [string]$reader["TypeCode"]
        DocName  = [string]$reader["DocName"]
        Category = [string]$reader["Category"]
        Bytes    = $reader["Bytes"]
        RptHash  = [string]$reader["RptHash"]
    })
}
$reader.Close()

$csvFile = $outFile -replace "\.bak$", ".csv"
$rows | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
Write-Host "Saved index of $($rows.Count) existing rows -> $csvFile" -ForegroundColor Green

# Binary dump (all templates with header)
$fs = [System.IO.File]::OpenWrite($outFile)
$bw = New-Object System.IO.BinaryWriter $fs
$bw.Write([byte[]]([byte]0x52,0x44,0x4F,0x43,0x42,0x4B,0x50,0x31)) # magic "RDOCBKP1"
$bw.Write([int32]$rows.Count)

$cmd2 = $conn.CreateCommand()
$cmd2.CommandText = Convert-DBSql "SELECT DocCode, Template FROM RDOC WHERE Template IS NOT NULL"
$reader = $cmd2.ExecuteReader()
$savedBinary = 0
while ($reader.Read()) {
    $code = [string]$reader["DocCode"]
    $blob = $reader["Template"]
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

Write-Host "Saved $savedBinary binary templates -> $outFile" -ForegroundColor Green
Write-Host ""
Write-Host "Backup complete:" -ForegroundColor Yellow
Write-Host "  Index : $csvFile"
Write-Host "  Binary: $outFile ($((Get-Item $outFile).Length) bytes)"
