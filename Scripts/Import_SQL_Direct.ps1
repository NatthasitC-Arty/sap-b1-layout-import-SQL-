# ============================================================
# Batch Import Crystal Layouts into SAP B1 (HANA, SQL Direct)
# Reads RPT_Import_Map.xlsx -> INSERT/UPDATE the company schema's RDOC table
# DocCode and TypeCode come straight from the mapping file (no
# ObjectType lookup) — dedup is by DocCode PK.
# ============================================================
param(
    [string]$Server      = "10.10.10.109:30015",
    [string]$CompanyDB   = "SBO_ENCONFUND_TRAINING",   # HANA schema
    [string]$DBUser      = "SYSTEM",
    [string]$DBPassword  = "",
    [string]$MapFile     = "$PSScriptRoot\..\Config\RPT_Import_Map.xlsx",
    [string]$RptRoot     = "C:\GitHub\Enconfund\FORM",
    [string]$LogFile     = "$PSScriptRoot\..\Import_SQL_Log.txt",
    [string]$Author      = "manager",
    [ValidateSet("Update","Skip")]
    [string]$OnDuplicate = "Update",
    [string]$FilterFileName = "",
    [switch]$UseFileNameAsDocName,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\DB-HANA.ps1"

$schemaQ = Get-DBQuoteIdent $CompanyDB

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Msg
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Read-MapExcel {
    # Excel schema (columns A..G):
    #   A No | B DocCode | C TypeCode | D RPT_FileName | E RPT_Folder | F LayoutName | G Note
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

        # Find sheet "RPT_MAP" rId
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
                DocCode      = $cells[2]
                TypeCode     = $cells[3]
                RPT_FileName = $cells[4]
                RPT_Folder   = $cells[5]
                LayoutName   = $cells[6]
                Note         = $cells[7]
            }
            if ($item.RPT_FileName -or $item.DocCode) { $list += $item }
        }
        return $list
    } finally {
        $zip.Dispose()
    }
}

Write-Log "=== Start HANA Direct Import ==="
Write-Log ("Server  : {0}" -f $Server)
Write-Log ("Schema  : {0}" -f $CompanyDB)
Write-Log ("MapFile : {0}" -f $MapFile)
Write-Log ("DryRun  : {0}" -f $DryRun)

$mapRows = Read-MapExcel -Path $MapFile
Write-Log ("Loaded {0} mapping rows" -f $mapRows.Count)

$conn = New-DBConnection -Server $Server -Database $CompanyDB -User $DBUser -Password $DBPassword
Write-Log ("Connected to HANA {0}" -f $conn.ServerVersion)

$ok = 0; $fail = 0; $skip = 0
foreach ($row in $mapRows) {
    if ($FilterFileName -and ($row.RPT_FileName -notlike "*$FilterFileName*")) { continue }
    $docCode  = [string]$row.DocCode
    $typeCode = [string]$row.TypeCode
    if ([string]::IsNullOrWhiteSpace($docCode))  { Write-Log ("SKIP no DocCode  : {0}" -f $row.RPT_FileName)  "WARN"; $skip++; continue }
    if ([string]::IsNullOrWhiteSpace($typeCode)) { Write-Log ("SKIP no TypeCode : {0}" -f $row.RPT_FileName)  "WARN"; $skip++; continue }

    $rptPath = Join-Path $RptRoot (Join-Path $row.RPT_Folder $row.RPT_FileName)
    if (-not (Test-Path $rptPath)) { Write-Log ("SKIP missing file: {0}" -f $rptPath) "WARN"; $skip++; continue }

    if ($UseFileNameAsDocName) {
        $layoutName = [System.IO.Path]::GetFileNameWithoutExtension($row.RPT_FileName)
    } else {
        $layoutName = if ([string]::IsNullOrWhiteSpace($row.LayoutName)) {
            [System.IO.Path]::GetFileNameWithoutExtension($row.RPT_FileName)
        } else {
            [string]$row.LayoutName
        }
    }

    try {
        $bytes = [System.IO.File]::ReadAllBytes($rptPath)
        $md5  = [System.Security.Cryptography.MD5]::Create()
        $hash = [BitConverter]::ToString($md5.ComputeHash($bytes)).Replace("-","")
        $md5.Dispose()

        # Lookup by DocCode PK; record Author so we can guard against system overwrite.
        $chk = $conn.CreateCommand()
        Set-DBCommandText $chk "SELECT Author, TypeCode FROM $schemaQ.RDOC WHERE DocCode=@d"
        Add-DBParam $chk "@d" $docCode
        $rdr = Invoke-DBReader $chk
        $existingAuthor   = $null
        $existingTypeCode = $null
        if ($rdr.Read()) {
            $existingAuthor   = [string]$rdr['Author']
            $existingTypeCode = [string]$rdr['TypeCode']
        }
        $rdr.Close()

        $action = ""
        if ($null -ne $existingAuthor) {
            if ($existingAuthor -eq 'System') {
                Write-Log ("SKIP system     [{0,3}] {1} (DocCode={2}) - Author=System, refusing to overwrite" -f $row.No, $row.RPT_FileName, $docCode) "WARN"
                $skip++; continue
            }
            if ($existingTypeCode -and $existingTypeCode -ne $typeCode) {
                Write-Log ("FAIL type drift [{0,3}] {1}: DB has TypeCode='{2}', Excel says '{3}' for DocCode={4}" -f $row.No, $row.RPT_FileName, $existingTypeCode, $typeCode, $docCode) "ERROR"
                $fail++; continue
            }
            if ($OnDuplicate -eq "Skip") {
                Write-Log ("SKIP exists     [{0,3}] {1} (DocCode={2})" -f $row.No, $row.RPT_FileName, $docCode) "WARN"
                $skip++; continue
            }
            $action = "UPDATE"
        } else {
            $action = "INSERT"
        }

        if ($DryRun) {
            Write-Log ("DRYRUN [{0,3}] {1} -> {2} DocCode={3} TypeCode={4} Bytes={5} Hash={6}" -f $row.No, $row.RPT_FileName, $action, $docCode, $typeCode, $bytes.Length, $hash.Substring(0,8))
            $ok++; continue
        }

        $cmd = $conn.CreateCommand()
        if ($action -eq "UPDATE") {
            $sql = "UPDATE $schemaQ.RDOC SET DocName=@DocName, Template=@Template, RptHash=@RptHash, UpdateDate=$DB_NOW WHERE DocCode=@DocCode"
            Set-DBCommandText $cmd $sql
            Add-DBParam   $cmd "@DocName"  $layoutName
            Add-BlobParam $cmd "@Template" $bytes
            Add-DBParam   $cmd "@RptHash"  $hash
            Add-DBParam   $cmd "@DocCode"  $docCode
        } else {
            $sql = "INSERT INTO $schemaQ.RDOC (DocCode,DocName,Author,Notes,Width,Height,LMargin,RMargin,TMargin,BMargin,CanChange,PaperSize,Oreint,GridSize,GridType,ShowGrid,SnapGrid,TypeCode,FrgnReport,CanSort,LeaderCode,FollowCode,SwapOnScrn,ScreenFont,ScrFOffset,SwpInEmail,EmailFont,EmFOffset,QString,QType,RobjCode,ExtName,ExtOnErr,NumRepArs,AlgnFooter,TimeFormat,DateFormat,NumCopy,GbiSupport,Use1stPrtr,Shading,Template,Category,CreateDate,Status,B1Version,CRVersion,Local,UseSysPref,ForMobile,TypeDetail,IsIMCE,CsUrl,RptHash) VALUES (@DocCode,@DocName,@Author,'',595,842,10,30,10,10,'Y','A4','P',10,'1','Y','Y',@TypeCode,'N','Y','','','N','Arial',-1,'N','Arial',-1,'','R',0,'','S',-1,'N','0','0',1,'N','N','Y',@Template,'C',$DB_NOW,'A','','','','Y','Y','','N','',@RptHash)"
            Set-DBCommandText $cmd $sql
            Add-DBParam   $cmd "@DocCode"  $docCode
            Add-DBParam   $cmd "@DocName"  $layoutName
            Add-DBParam   $cmd "@Author"   $Author
            Add-DBParam   $cmd "@TypeCode" $typeCode
            Add-BlobParam $cmd "@Template" $bytes
            Add-DBParam   $cmd "@RptHash"  $hash
        }

        [void](Invoke-DBNonQuery $cmd)
        Write-Log ("{0} [{1,3}] {2} -> DocCode={3} ({4} bytes, {5})" -f $action, $row.No, $row.RPT_FileName, $docCode, $bytes.Length, $layoutName)
        $ok++
    } catch {
        Write-Log ("FAIL [{0,3}] {1}: {2}" -f $row.No, $row.RPT_FileName, $_.Exception.Message) "ERROR"
        $fail++
    }
}
$conn.Close()
Write-Log "=== Summary: OK=$ok FAIL=$fail SKIP=$skip ==="
