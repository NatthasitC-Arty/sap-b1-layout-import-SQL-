# ============================================================
# Delete ALL layouts in RDOC EXCEPT system defaults.
# "System" = rows where Author = 'System' (configurable via -SystemAuthor).
# Also deletes child rows in RITM/RDC1/RCON for the deleted DocCodes
# and clears DFLT_PRNTING orphans. All work is wrapped in a transaction.
# Works on MSSQL and HANA via the DB plugin.
# ============================================================
param(
    [Parameter(Mandatory=$true)][string]$Server,
    [Parameter(Mandatory=$true)][string]$CompanyDB,
    [string]$DBUser     = "sa",
    [Parameter(Mandatory=$true)][string]$DBPassword,
    [string]$SystemAuthor = "System",
    [ValidateSet("MSSQL","HANA")]
    [string]$DBEngine   = "MSSQL",
    [switch]$DryRun,
    [switch]$Force
)

# Load DB plugin
. "$PSScriptRoot\DB-$DBEngine.ps1"

Write-Host "=== Delete non-system layouts from RDOC ===" -ForegroundColor Cyan
Write-Host "Server      : $Server"
Write-Host "CompanyDB   : $CompanyDB"
Write-Host "DBEngine    : $DBEngine"
Write-Host "Keep Author : '$SystemAuthor'"
Write-Host ""

try {
    $conn = New-DBConnection -Server $Server -Database $CompanyDB -User $DBUser -Password $DBPassword
} catch {
    Write-Host "ERROR connecting: $($_.Exception.Message)" -ForegroundColor Red
    return
}

$tCmd = $conn.CreateCommand()
$tCmd.CommandText = Convert-DBSql "SELECT COUNT(*) FROM RDOC"
$total = $tCmd.ExecuteScalar()
Write-Host ("RDOC total rows           : {0}" -f $total) -ForegroundColor Gray

$sysCmd = $conn.CreateCommand()
$sysCmd.CommandText = Convert-DBSql "SELECT COUNT(*) FROM RDOC WHERE Author=${DB_PARAM}a"
Add-DBParam $sysCmd "${DB_PARAM}a" $SystemAuthor
$sysCount = $sysCmd.ExecuteScalar()
Write-Host ("Rows with Author='$SystemAuthor' (KEEP) : {0}" -f $sysCount) -ForegroundColor Green

$delCmd = $conn.CreateCommand()
$delCmd.CommandText = Convert-DBSql "SELECT COUNT(*) FROM RDOC WHERE Author<>${DB_PARAM}a OR Author IS NULL"
Add-DBParam $delCmd "${DB_PARAM}a" $SystemAuthor
$delCount = $delCmd.ExecuteScalar()
Write-Host ("Rows to DELETE            : {0}" -f $delCount) -ForegroundColor Yellow
Write-Host ""

if ($delCount -eq 0) {
    Write-Host "Nothing to delete." -ForegroundColor Green
    $conn.Close()
    return
}

# Preview: group by Author + TypeCode (manual reader loop, engine-agnostic)
$prev = $conn.CreateCommand()
$prev.CommandText = Convert-DBSql "SELECT Author, TypeCode, COUNT(*) AS Cnt FROM RDOC WHERE Author<>${DB_PARAM}a OR Author IS NULL GROUP BY Author, TypeCode ORDER BY Author, TypeCode"
Add-DBParam $prev "${DB_PARAM}a" $SystemAuthor
$rdr = $prev.ExecuteReader()
$prevRows = New-Object System.Collections.ArrayList
while ($rdr.Read()) {
    [void]$prevRows.Add([PSCustomObject]@{
        Author   = if ($rdr.IsDBNull(0)) { '<NULL>' } else { [string]$rdr[0] }
        TypeCode = [string]$rdr[1]
        Cnt      = $rdr[2]
    })
}
$rdr.Close()
Write-Host "=== Preview (grouped by Author + TypeCode) ===" -ForegroundColor Yellow
$prevRows | Format-Table Author,TypeCode,Cnt -AutoSize

if ($DryRun) {
    Write-Host "DryRun mode - no changes made." -ForegroundColor Cyan
    $conn.Close()
    return
}

if (-not $Force) {
    Write-Host "WARNING: This will permanently delete $delCount rows from RDOC." -ForegroundColor Red
    $ans = Read-Host "Type 'yes' to confirm"
    if ($ans -ne "yes") {
        Write-Host "Cancelled." -ForegroundColor Yellow
        $conn.Close()
        return
    }
}

$tran = $conn.BeginTransaction()
try {
    # Delete children first (subquery against RDOC works while parent rows still exist)
    foreach ($tbl in @('RITM','RDC1','RCON')) {
        try {
            $c = $conn.CreateCommand()
            $c.Transaction = $tran
            $c.CommandText = Convert-DBSql "DELETE FROM dbo.$tbl WHERE DocCode IN (SELECT DocCode FROM RDOC WHERE Author<>${DB_PARAM}a OR Author IS NULL)"
            Add-DBParam $c "${DB_PARAM}a" $SystemAuthor
            $c.CommandTimeout = 300
            $cn = $c.ExecuteNonQuery()
            Write-Host ("  {0,-4}: deleted {1,5} child rows" -f $tbl, $cn) -ForegroundColor DarkYellow
        } catch {
            Write-Host ("  {0,-4}: skipped ({1})" -f $tbl, $_.Exception.Message) -ForegroundColor DarkGray
        }
    }

    # Delete RDOC parent rows
    $exec = $conn.CreateCommand()
    $exec.Transaction = $tran
    $exec.CommandText = Convert-DBSql "DELETE FROM RDOC WHERE Author<>${DB_PARAM}a OR Author IS NULL"
    Add-DBParam $exec "${DB_PARAM}a" $SystemAuthor
    $exec.CommandTimeout = 300
    $n = $exec.ExecuteNonQuery()
    Write-Host ("  RDOC: deleted {0,5} rows" -f $n) -ForegroundColor Green

    # Clean DFLT_PRNTING orphans (table may not exist on older versions)
    try {
        $d = $conn.CreateCommand()
        $d.Transaction = $tran
        $d.CommandText = Convert-DBSql "DELETE FROM dbo.DFLT_PRNTING WHERE DocCode IS NOT NULL AND DocCode<>'' AND DocCode NOT IN (SELECT DocCode FROM RDOC)"
        $d.CommandTimeout = 300
        $dn = $d.ExecuteNonQuery()
        Write-Host ("  DFLT_PRNTING orphans cleaned: {0,5} rows" -f $dn) -ForegroundColor DarkYellow
    } catch {
        Write-Host ("  DFLT_PRNTING: skipped ({0})" -f $_.Exception.Message) -ForegroundColor DarkGray
    }

    $tran.Commit()
    Write-Host ""
    Write-Host "=== Summary: Deleted=$n / Kept(system)=$sysCount ===" -ForegroundColor Cyan
} catch {
    $tran.Rollback()
    Write-Host ""
    Write-Host ("[ROLLBACK] {0}" -f $_.Exception.Message) -ForegroundColor Red
    throw
} finally {
    $conn.Close()
}
