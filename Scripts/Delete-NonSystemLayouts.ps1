# ============================================================
# Delete ALL layouts in RDOC EXCEPT system defaults (HANA edition).
# "System" = rows where Author = 'System' (configurable via -SystemAuthor).
# Also deletes child rows in RITM/RDC1/RCON for the deleted DocCodes
# and clears DFLT_PRNTING orphans. All work is wrapped in a transaction.
# ============================================================
param(
    [Parameter(Mandatory=$true)][string]$Server,
    [Parameter(Mandatory=$true)][string]$CompanyDB,    # HANA schema
    [string]$DBUser     = "SYSTEM",
    [Parameter(Mandatory=$true)][string]$DBPassword,
    [string]$SystemAuthor = "System",
    [switch]$DryRun,
    [switch]$Force
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\DB-HANA.ps1"

$schemaQ = Get-DBQuoteIdent $CompanyDB

Write-Host "=== Delete non-system layouts from RDOC (HANA) ===" -ForegroundColor Cyan
Write-Host "Server      : $Server"
Write-Host "Schema      : $CompanyDB"
Write-Host "Keep Author : '$SystemAuthor'"
Write-Host ""

try {
    $conn = New-DBConnection -Server $Server -Database $CompanyDB -User $DBUser -Password $DBPassword
} catch {
    Write-Host "ERROR connecting: $($_.Exception.Message)" -ForegroundColor Red
    return
}

$cmd = $conn.CreateCommand()
Set-DBCommandText $cmd "SELECT COUNT(*) FROM $schemaQ.RDOC"
$total = Invoke-DBScalar $cmd
Write-Host ("RDOC total rows                : {0}" -f $total) -ForegroundColor Gray

$cmdSys = $conn.CreateCommand()
Set-DBCommandText $cmdSys "SELECT COUNT(*) FROM $schemaQ.RDOC WHERE Author=@a"
Add-DBParam $cmdSys "@a" $SystemAuthor
$sysCount = Invoke-DBScalar $cmdSys
Write-Host ("Rows with Author='$SystemAuthor' (KEEP) : {0}" -f $sysCount) -ForegroundColor Green

$cmdDel = $conn.CreateCommand()
Set-DBCommandText $cmdDel "SELECT COUNT(*) FROM $schemaQ.RDOC WHERE Author<>@a OR Author IS NULL"
Add-DBParam $cmdDel "@a" $SystemAuthor
$delCount = Invoke-DBScalar $cmdDel
Write-Host ("Rows to DELETE                 : {0}" -f $delCount) -ForegroundColor Yellow
Write-Host ""

if ($delCount -eq 0) {
    Write-Host "Nothing to delete." -ForegroundColor Green
    $conn.Close()
    return
}

# Preview: group by Author + TypeCode
$cmdPrev = $conn.CreateCommand()
Set-DBCommandText $cmdPrev @"
SELECT Author, TypeCode, COUNT(*) AS Cnt
FROM $schemaQ.RDOC
WHERE Author<>@a OR Author IS NULL
GROUP BY Author, TypeCode
ORDER BY Author, TypeCode
"@
Add-DBParam $cmdPrev "@a" $SystemAuthor
$rdr = Invoke-DBReader $cmdPrev
$preview = New-Object System.Collections.Generic.List[object]
while ($rdr.Read()) {
    $preview.Add([PSCustomObject]@{
        Author   = [string]$rdr['Author']
        TypeCode = [string]$rdr['TypeCode']
        Cnt      = [int]$rdr['Cnt']
    })
}
$rdr.Close()
Write-Host "=== Preview (grouped by Author + TypeCode) ===" -ForegroundColor Yellow
$preview | Format-Table Author, TypeCode, Cnt -AutoSize

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
    # Delete child rows first (subqueries against RDOC work while parent rows still exist)
    foreach ($tbl in @('RITM','RDC1','RCON')) {
        try {
            $c = $conn.CreateCommand()
            $c.Transaction = $tran
            $c.CommandTimeout = 300
            Set-DBCommandText $c "DELETE FROM $schemaQ.`"$tbl`" WHERE DocCode IN (SELECT DocCode FROM $schemaQ.RDOC WHERE Author<>@a OR Author IS NULL)"
            Add-DBParam $c "@a" $SystemAuthor
            $cn = Invoke-DBNonQuery $c
            Write-Host ("  {0,-4}: deleted {1,5} child rows" -f $tbl, $cn) -ForegroundColor DarkYellow
        } catch {
            Write-Host ("  {0,-4}: skipped ({1})" -f $tbl, $_.Exception.Message) -ForegroundColor DarkGray
        }
    }

    # Delete RDOC parent rows
    $exec = $conn.CreateCommand()
    $exec.Transaction = $tran
    $exec.CommandTimeout = 300
    Set-DBCommandText $exec "DELETE FROM $schemaQ.RDOC WHERE Author<>@a OR Author IS NULL"
    Add-DBParam $exec "@a" $SystemAuthor
    $n = Invoke-DBNonQuery $exec
    Write-Host ("  RDOC: deleted {0,5} rows" -f $n) -ForegroundColor Green

    # Clean DFLT_PRNTING orphans
    try {
        $d = $conn.CreateCommand()
        $d.Transaction = $tran
        $d.CommandTimeout = 300
        Set-DBCommandText $d "DELETE FROM $schemaQ.DFLT_PRNTING WHERE DocCode IS NOT NULL AND DocCode<>'' AND DocCode NOT IN (SELECT DocCode FROM $schemaQ.RDOC)"
        $dn = Invoke-DBNonQuery $d
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
