# ============================================================
# Test DB connection to SAP B1 Company DB
# Verifies you can read RDOC table before running import.
# Works on MSSQL and HANA via the DB plugin.
# ============================================================
param(
    [string]$Server     = "SLD-C072",
    [string]$CompanyDB  = "SBO_SDA",
    [string]$DBUser     = "sa",
    [string]$DBPassword = "1q2w3e4r",
    [ValidateSet("MSSQL","HANA")]
    [string]$DBEngine   = "MSSQL"
)

$ErrorActionPreference = "Stop"

# Load DB plugin
. "$PSScriptRoot\DB-$DBEngine.ps1"

# Resolve host/port for the TCP test
$hostName = $Server
$port = if ($DBEngine -eq "HANA") { 30015 } else { 1433 }
if ($Server -match '^([^:,]+)[:,](\d+)$') {
    $hostName = $matches[1]
    $port = [int]$matches[2]
}

Write-Host "[1/4] Pinging $hostName ..." -ForegroundColor Cyan
$ping = Test-Connection -ComputerName $hostName -Count 2 -Quiet -ErrorAction SilentlyContinue
Write-Host "      Ping: $(if($ping){'OK'}else{'FAIL (host not reachable)'})" -ForegroundColor $(if($ping){'Green'}else{'Red'})

Write-Host "[2/4] Testing TCP port $port on $hostName ..." -ForegroundColor Cyan
$tcp = Test-NetConnection -ComputerName $hostName -Port $port -WarningAction SilentlyContinue
Write-Host "      TCP $port`: $(if($tcp.TcpTestSucceeded){'OPEN'}else{'CLOSED/BLOCKED'})" -ForegroundColor $(if($tcp.TcpTestSucceeded){'Green'}else{'Red'})

Write-Host "[3/4] Testing $DBEngine connection to $CompanyDB ..." -ForegroundColor Cyan
try {
    $conn = New-DBConnection -Server $Server -Database $CompanyDB -User $DBUser -Password $DBPassword
    Write-Host "      $DBEngine Login: OK (server $($conn.ServerVersion))" -ForegroundColor Green

    Write-Host "[4/4] Counting layouts in RDOC ..." -ForegroundColor Cyan
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = Convert-DBSql "SELECT COUNT(*) AS Total, SUM(CASE WHEN Category='C' THEN 1 ELSE 0 END) AS Crystal, SUM(CASE WHEN Author='SDA' THEN 1 ELSE 0 END) AS SDA FROM RDOC"
    $rdr = $cmd.ExecuteReader()
    if ($rdr.Read()) {
        Write-Host "      Total layouts : $($rdr['Total'])" -ForegroundColor Green
        Write-Host "      Crystal (C)   : $($rdr['Crystal'])" -ForegroundColor Green
        Write-Host "      Author=SDA    : $($rdr['SDA']) (imported by us)" -ForegroundColor Green
    }
    $rdr.Close()
    $conn.Close()
    Write-Host ""
    Write-Host "READY TO IMPORT" -ForegroundColor Green
} catch {
    Write-Host "      $DBEngine FAIL: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshoot:" -ForegroundColor Yellow
    if ($DBEngine -eq "HANA") {
        Write-Host "  - 'authentication failed' -> wrong DBUser/DBPassword (HANA is case-sensitive)"
        Write-Host "  - 'invalid schema name' -> wrong CompanyDB; HANA uses the COMPANY SCHEMA, e.g. SBO_XXX"
        Write-Host "  - 'connect failed'      -> wrong Server (host:port). Tenant DB port is 3<NN>15"
        Write-Host "  - 'Sap.Data.Hana not found' -> install SAP HANA Client (hdbclient)"
    } else {
        Write-Host "  - 'Login failed for user' -> wrong DBUser/DBPassword"
        Write-Host "  - 'Cannot open database X' -> wrong CompanyDB name"
        Write-Host "  - 'A network-related error' -> wrong Server name or SQL service down"
    }
}
