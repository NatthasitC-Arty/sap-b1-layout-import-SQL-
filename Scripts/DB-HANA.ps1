# ============================================================
# DB Plugin: SAP HANA (B1 on HANA)
# Dot-sourced by main scripts via: . "$PSScriptRoot\DB-HANA.ps1"
# Requires SAP HANA Client (Sap.Data.Hana.v4.5.dll) on this machine.
# ============================================================

$DB_PARAM    = ":"                  # HANA named-parameter prefix
$DB_NOW      = "CURRENT_TIMESTAMP"  # current timestamp
$DB_ISNUM    = "ISNUMERIC"          # Convert-DBSql rewrites ISNUMERIC(x)=1 to TRY_CAST(x AS INT) IS NOT NULL

$script:HanaDllPath = "C:\Program Files\SAP\hdbclient\ado.net\v4.5\Sap.Data.Hana.v4.5.dll"

function Initialize-HanaDriver {
    $loaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Sap.Data.Hana.v4.5' }
    if ($loaded) { return }
    if (-not (Test-Path $script:HanaDllPath)) {
        throw "SAP HANA Client not found at '$script:HanaDllPath'. Install hdbclient (e.g. from SAP HANA Client setup) and retry."
    }
    # LoadFrom (not Add-Type -Path) because Add-Type validates every exported type
    # via GetTypes(), which throws ReflectionTypeLoadException when the optional
    # EntityFramework 6.0 reference (pulled in by HanaEF*) is absent. HanaConnection
    # itself loads and runs fine without EF -- we just bypass the eager validation.
    [void][System.Reflection.Assembly]::LoadFrom($script:HanaDllPath)
}

function New-DBConnection {
    param(
        [string]$Server,    # host:port (e.g. 10.10.10.109:30015)
        [string]$Database,  # HANA: company SCHEMA name (e.g. SBO_ENCONFUND_TRAINING)
        [string]$User,
        [string]$Password,
        [int]$Timeout = 10
    )
    Initialize-HanaDriver
    $cs = "Server=$Server;UID=$User;PWD=$Password;CurrentSchema=$Database;Pooling=true;Connect Timeout=$Timeout;"
    $conn = New-Object Sap.Data.Hana.HanaConnection $cs
    $conn.Open()
    return $conn
}

function Add-BlobParam {
    param($Command, [string]$Name, [byte[]]$Bytes)
    $clean = $Name -replace '^[:@]', ''
    $p = $Command.Parameters.Add($clean, [Sap.Data.Hana.HanaDbType]::Blob)
    $p.Direction = [System.Data.ParameterDirection]::Input
    $p.Value = $Bytes
}

function Add-DBParam {
    param($Command, [string]$Name, $Value)
    $clean = $Name -replace '^[:@]', ''
    [void]$Command.Parameters.AddWithValue($clean, $Value)
}

# ------------------------------------------------------------
# SQL dialect translation: MSSQL -> HANA
# Scripts write SQL in MSSQL flavour; Convert-DBSql rewrites it
# for HANA. Only what this tool actually emits is handled.
# ------------------------------------------------------------
function Convert-DBSql {
    param([string]$Sql)
    if ([string]::IsNullOrEmpty($Sql)) { return $Sql }
    $s = $Sql

    # 1. Strip identifier brackets:  [name] -> name   ([dbo], [SBO_X], etc.)
    $s = [regex]::Replace($s, '\[([^\]]+)\]', '$1')

    # 2. Strip dbo. schema prefix (HANA uses CurrentSchema from the connection)
    $s = [regex]::Replace($s, '\bdbo\.', '', 'IgnoreCase')

    # 3. Function renames
    $s = [regex]::Replace($s, '\bGETDATE\s*\(\s*\)', 'CURRENT_TIMESTAMP', 'IgnoreCase')
    $s = [regex]::Replace($s, '\bLEN\s*\(',         'LENGTH(',           'IgnoreCase')
    $s = [regex]::Replace($s, '\bDATALENGTH\s*\(',  'LENGTH(',           'IgnoreCase')
    $s = [regex]::Replace($s, '\bISNULL\s*\(',      'COALESCE(',         'IgnoreCase')

    # 4. ISNUMERIC(expr)=1  ->  TRY_CAST(expr AS INT) IS NOT NULL
    #    Hand-rolled because expr can contain nested parens.
    $s = ConvertTo-HanaIsNumeric $s

    return $s
}

function ConvertTo-HanaIsNumeric {
    param([string]$Sql)
    $sb = New-Object System.Text.StringBuilder
    $pattern = [regex]'(?i)ISNUMERIC\s*\('
    $i = 0
    while ($i -lt $Sql.Length) {
        $m = $pattern.Match($Sql, $i)
        if (-not $m.Success) {
            [void]$sb.Append($Sql.Substring($i))
            break
        }
        [void]$sb.Append($Sql.Substring($i, $m.Index - $i))
        $argStart = $m.Index + $m.Length
        $depth = 1
        $k = $argStart
        while ($k -lt $Sql.Length -and $depth -gt 0) {
            $ch = $Sql[$k]
            if ($ch -eq '(') { $depth++ }
            elseif ($ch -eq ')') { $depth--; if ($depth -eq 0) { break } }
            $k++
        }
        if ($depth -ne 0) {
            [void]$sb.Append($Sql.Substring($m.Index))
            break
        }
        $arg = $Sql.Substring($argStart, $k - $argStart)
        $after = $k + 1
        $tail = [regex]::Match($Sql.Substring($after), '^\s*=\s*1')
        if ($tail.Success) {
            [void]$sb.Append("TRY_CAST($arg AS INT) IS NOT NULL")
            $i = $after + $tail.Length
        } else {
            [void]$sb.Append("TRY_CAST($arg AS INT) IS NOT NULL")
            $i = $after
        }
    }
    return $sb.ToString()
}
