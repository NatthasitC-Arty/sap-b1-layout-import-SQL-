# ============================================================
# Test connection to SAP HANA Company tenant DB
# Verifies you can read RDOC table before running import
# ============================================================
param(
    [string]$Server     = "10.10.10.109:30015",
    [string]$CompanyDB  = "SBO_ENCONFUND_TRAINING",   # HANA schema
    [string]$DBUser     = "SYSTEM",
    [string]$DBPassword = ""
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\DB-HANA.ps1"

$schemaQ = Get-DBQuoteIdent $CompanyDB

# ---------- 1. Reachability ----------
$hostOnly = ($Server -split ':')[0]
$port     = if ($Server -match ':(\d+)$') { [int]$matches[1] } else { 30015 }

Write-Host "[1/4] Pinging $hostOnly ..." -ForegroundColor Cyan
$ping = Test-Connection -ComputerName $hostOnly -Count 2 -Quiet -ErrorAction SilentlyContinue
Write-Host "      Ping: $(if($ping){'OK'}else{'FAIL (host not reachable)'})" -ForegroundColor $(if($ping){'Green'}else{'Red'})

Write-Host "[2/4] Testing TCP $port on $hostOnly ..." -ForegroundColor Cyan
$tcp = Test-NetConnection -ComputerName $hostOnly -Port $port -WarningAction SilentlyContinue
Write-Host "      TCP $port`: $(if($tcp.TcpTestSucceeded){'OPEN'}else{'CLOSED/BLOCKED'})" -ForegroundColor $(if($tcp.TcpTestSucceeded){'Green'}else{'Red'})

# ---------- 2. HANA login ----------
Write-Host "[3/4] Connecting to HANA schema $CompanyDB ..." -ForegroundColor Cyan
try {
    $conn = New-DBConnection -Server $Server -Database $CompanyDB -User $DBUser -Password $DBPassword -Timeout 15
    Write-Host ("      HANA Login: OK (server {0})" -f $conn.ServerVersion) -ForegroundColor Green
} catch {
    Write-Host ("      HANA FAIL: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshoot:" -ForegroundColor Yellow
    Write-Host "  - 'authentication failed'  -> wrong DBUser/DBPassword"
    Write-Host "  - 'cannot find schema'     -> wrong CompanyDB (schema name) - check SELECT SCHEMA_NAME FROM SYS.SCHEMAS"
    Write-Host "  - 'cannot connect'         -> wrong Server:port or HANA service down"
    Write-Host "  - 'Sap.Data.Hana not found' -> install SAP HANA Client from SAP Marketplace"
    return
}

# ---------- 3. RDOC counts ----------
Write-Host "[4/4] Counting layouts in $schemaQ.RDOC ..." -ForegroundColor Cyan
try {
    $cmd = $conn.CreateCommand()
    Set-DBCommandText $cmd "SELECT COUNT(*) AS Total, SUM(CASE WHEN Category='C' THEN 1 ELSE 0 END) AS Crystal, SUM(CASE WHEN Author='SDA' THEN 1 ELSE 0 END) AS SDA FROM $schemaQ.RDOC"
    $rdr = Invoke-DBReader $cmd
    if ($rdr.Read()) {
        Write-Host ("      Total layouts : {0}" -f $rdr['Total']) -ForegroundColor Green
        Write-Host ("      Crystal (C)   : {0}" -f $rdr['Crystal']) -ForegroundColor Green
        Write-Host ("      Author=SDA    : {0}" -f $rdr['SDA']) -ForegroundColor Green
    }
    $rdr.Close()
    Write-Host ""
    Write-Host "READY TO IMPORT" -ForegroundColor Green
} catch {
    Write-Host ("      Query FAIL: {0}" -f $_.Exception.Message) -ForegroundColor Red
} finally {
    $conn.Close()
}
