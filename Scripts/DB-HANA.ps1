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

    # 5. Quote PascalCase identifiers ("DocCode", "Author", etc.) for HANA.
    #    SAP B1 stores its column names as mixed-case quoted identifiers; an
    #    unquoted reference is upcased by HANA (Category -> CATEGORY) and fails
    #    with "invalid column name". Tables (RDOC, RITM, ...) and keywords are
    #    all-uppercase so they don't match the PascalCase heuristic.
    $s = Add-HanaIdentifierQuotes $s

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

# ------------------------------------------------------------
# Wrap PascalCase identifiers in double quotes for HANA.
# Heuristic: an identifier that starts with [A-Z] AND contains at least
# one [a-z] is treated as a B1 mixed-case column/alias (DocCode, Author,
# Total, ...) and gets quoted. ALL_CAPS tokens (RDOC, SELECT, FROM,
# DFLT_PRNTING, LENGTH, COALESCE, ...) don't match and stay unquoted.
# Skips: contents of '...' string literals, contents of "..." already-
# quoted identifiers, and tokens immediately following ':' or '@' (named
# parameters like :DocCode).
# ------------------------------------------------------------
function Add-HanaIdentifierQuotes {
    param([string]$Sql)
    $sb = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $Sql.Length) {
        $ch = $Sql[$i]

        # Pass through string literal '...' verbatim
        if ($ch -eq "'") {
            [void]$sb.Append($ch); $i++
            while ($i -lt $Sql.Length) {
                $c2 = $Sql[$i]
                [void]$sb.Append($c2); $i++
                if ($c2 -eq "'") {
                    if ($i -lt $Sql.Length -and $Sql[$i] -eq "'") {
                        [void]$sb.Append($Sql[$i]); $i++  # escaped ''
                        continue
                    }
                    break
                }
            }
            continue
        }

        # Pass through already-quoted identifier "..." verbatim
        if ($ch -eq '"') {
            [void]$sb.Append($ch); $i++
            while ($i -lt $Sql.Length) {
                $c2 = $Sql[$i]; [void]$sb.Append($c2); $i++
                if ($c2 -eq '"') { break }
            }
            continue
        }

        # Identifier start
        if (($ch -ge 'A' -and $ch -le 'Z') -or ($ch -ge 'a' -and $ch -le 'z') -or $ch -eq '_') {
            $j = $i + 1
            while ($j -lt $Sql.Length) {
                $cj = $Sql[$j]
                if ((($cj -ge 'A') -and ($cj -le 'Z')) -or
                    (($cj -ge 'a') -and ($cj -le 'z')) -or
                    (($cj -ge '0') -and ($cj -le '9')) -or
                    $cj -eq '_') { $j++ } else { break }
            }
            $token = $Sql.Substring($i, $j - $i)
            $prev = if ($sb.Length -gt 0) { $sb[$sb.Length - 1] } else { ' ' }
            $isParam = ($prev -eq ':' -or $prev -eq '@')
            $startsUpper = ($token[0] -ge 'A' -and $token[0] -le 'Z')
            $hasLower = $false
            for ($k = 0; $k -lt $token.Length; $k++) {
                if ($token[$k] -ge 'a' -and $token[$k] -le 'z') { $hasLower = $true; break }
            }
            if ((-not $isParam) -and $startsUpper -and $hasLower) {
                [void]$sb.Append('"').Append($token).Append('"')
            } else {
                [void]$sb.Append($token)
            }
            $i = $j
            continue
        }

        [void]$sb.Append($ch); $i++
    }
    return $sb.ToString()
}
