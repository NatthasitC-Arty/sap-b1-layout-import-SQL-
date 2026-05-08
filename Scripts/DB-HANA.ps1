# ============================================================
# DB Plugin: SAP HANA (via Sap.Data.Hana .NET data provider)
# Dot-sourced by main scripts via: . "$PSScriptRoot\DB-HANA.ps1"
#
# CONTRACT
#   $DB_PARAM     parameter prefix for SQL strings ('@' so call sites stay
#                 readable; the plugin re-writes to HANA's positional '?'
#                 inside Set-DBCommandText)
#   $DB_NOW       expression for current timestamp (CURRENT_TIMESTAMP)
#   $DB_ISNUM     name-only ISNUMERIC stand-in (HANA has no ISNUMERIC; this
#                 is left empty and call sites that need it use a regex)
#   $DB_QUOTE     identifier-quote helper (HANA = double-quote)
#
# CALL SITE PATTERN (REQUIRED — do not assign $cmd.CommandText directly):
#
#   $cmd = $conn.CreateCommand()
#   Set-DBCommandText $cmd "SELECT * FROM RDOC WHERE DocCode=@d"
#   Add-DBParam   $cmd "@d" $docCode
#   Add-BlobParam $cmd "@t" $bytes        # for byte[] / BLOB columns
#   $rows = Invoke-DBNonQuery $cmd        # or Invoke-DBScalar / Invoke-DBReader
#
# WHY THIS PATTERN
#   HANA's HanaCommand only supports positional '?' parameters. Set-DBCommandText
#   parses '@name' tokens and rewrites them to '?', stashing the order on the
#   command. Add-DBParam looks up the position by name. Invoke-DB* binds the
#   slots in order before executing. This lets call sites use named placeholders
#   for readability without depending on driver-side named binding.
# ============================================================

$DB_PARAM = "@"
$DB_NOW   = "CURRENT_TIMESTAMP"
$DB_ISNUM = ""

$script:HanaAssemblyLoaded = $false

function Initialize-HanaDriver {
    if ($script:HanaAssemblyLoaded) { return }
    $candidates = @(
        "C:\Program Files\SAP\hdbclient\dotnetv45\Sap.Data.Hana.v4.5.dll",
        "C:\Program Files\SAP\hdbclient\Sap.Data.Hana.v4.5.dll",
        "C:\Program Files\SAP\hdbclient\dotnetcore\Sap.Data.Hana.Core.v2.1.dll",
        "C:\Program Files\sap\hdbclient\dotnetv45\Sap.Data.Hana.v4.5.dll",
        "C:\Program Files\sap\hdbclient\Sap.Data.Hana.v4.5.dll"
    )
    foreach ($p in $candidates) {
        if (Test-Path $p) {
            try {
                Add-Type -Path $p -ErrorAction Stop
                $script:HanaAssemblyLoaded = $true
                Write-Verbose "Loaded HANA provider: $p"
                return
            } catch {
                Write-Verbose "Failed to load $p : $($_.Exception.Message)"
            }
        }
    }
    # Last resort: try GAC partial load
    try {
        $a = [System.Reflection.Assembly]::LoadWithPartialName("Sap.Data.Hana")
        if ($a) { $script:HanaAssemblyLoaded = $true; return }
    } catch {}
    throw "SAP HANA .NET data provider (Sap.Data.Hana.dll) not found. Install SAP HANA Client from SAP Marketplace."
}

function New-DBConnection {
    param(
        [string]$Server,         # "host:30015" or just "host" (port defaults to 30015)
        [string]$Database,       # interpreted as HANA CurrentSchema
        [string]$User,
        [string]$Password,
        [int]$Timeout = 15
    )
    Initialize-HanaDriver
    if ($Server -notmatch ":") { $Server = "${Server}:30015" }
    $cs = "Server=$Server;UserID=$User;Password=$Password;CurrentSchema=$Database;CommunicationTimeout=$($Timeout * 1000);"
    $conn = New-Object Sap.Data.Hana.HanaConnection $cs
    $conn.Open()
    return $conn
}

function Convert-DBSql {
    # No-op for symmetry with DB-MSSQL contract. SQL translation happens
    # inside Set-DBCommandText so the parameter rewrite and the order
    # tracking stay in one place.
    param([string]$Sql)
    return $Sql
}

function Set-DBCommandText {
    # Parse '@name' tokens, replace with '?', stash the order on the command.
    # Same name appearing N times in SQL produces N positional binds.
    param($Command, [string]$Sql)
    $orderList = New-Object System.Collections.Generic.List[string]
    $regex = [regex]'@([A-Za-z_][A-Za-z0-9_]*)'
    $newSql = $regex.Replace($Sql, {
        param($m)
        $orderList.Add($m.Groups[1].Value)
        return "?"
    })
    $values = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $blobs  = New-Object 'System.Collections.Generic.HashSet[string]'
    $Command | Add-Member -NotePropertyName "_HanaOrder"  -NotePropertyValue $orderList -Force
    $Command | Add-Member -NotePropertyName "_HanaValues" -NotePropertyValue $values    -Force
    $Command | Add-Member -NotePropertyName "_HanaBlobs"  -NotePropertyValue $blobs     -Force
    $Command.CommandText = $newSql
}

function Add-DBParam {
    param($Command, [string]$Name, $Value)
    $n = $Name.TrimStart('@').TrimStart(':')
    if (-not $Command._HanaValues) {
        throw "Set-DBCommandText must be called before Add-DBParam (cmd has no _HanaValues)"
    }
    $v = $Value
    if ($null -eq $v) { $v = [System.DBNull]::Value }
    $Command._HanaValues[$n] = $v
}

function Add-BlobParam {
    param($Command, [string]$Name, [byte[]]$Bytes)
    $n = $Name.TrimStart('@').TrimStart(':')
    if (-not $Command._HanaValues) {
        throw "Set-DBCommandText must be called before Add-BlobParam"
    }
    $Command._HanaValues[$n] = $Bytes
    [void]$Command._HanaBlobs.Add($n)
}

function Submit-DBParams {
    # Build positional Hana parameters from _HanaOrder + _HanaValues.
    # Re-bindable: clears existing parameters first.
    param($Command)
    $Command.Parameters.Clear()
    if (-not $Command._HanaOrder) { return }
    foreach ($name in $Command._HanaOrder) {
        if (-not $Command._HanaValues.ContainsKey($name)) {
            throw "Parameter '@$name' referenced in SQL but never bound. Add-DBParam/Add-BlobParam was not called for it."
        }
        $val = $Command._HanaValues[$name]
        $isBlob = $Command._HanaBlobs.Contains($name)
        $p = New-Object Sap.Data.Hana.HanaParameter
        if ($isBlob) {
            $p.HanaDbType = [Sap.Data.Hana.HanaDbType]::Blob
        }
        $p.Value = $val
        [void]$Command.Parameters.Add($p)
    }
}

function Invoke-DBNonQuery {
    param($Command)
    Submit-DBParams $Command
    return $Command.ExecuteNonQuery()
}

function Invoke-DBScalar {
    param($Command)
    Submit-DBParams $Command
    return $Command.ExecuteScalar()
}

function Invoke-DBReader {
    param($Command)
    Submit-DBParams $Command
    return $Command.ExecuteReader()
}

function Get-DBQuoteIdent {
    # HANA identifiers — wrap in double quotes. Used to schema-qualify tables:
    # FROM "SBO_ENCONFUND_TRAINING"."RDOC"
    param([string]$Name)
    return '"' + ($Name -replace '"','""') + '"'
}
