<#
.SYNOPSIS
    Captures a tablespace usage snapshot from the Production Oracle DB.

.DESCRIPTION
    Queries dba_data_files and dba_free_space on the Production database and
    writes the result as TablespaceSnapshot.json to the ExportBaseDirectory.
    This file is copied to Stage by Start-DataTransfer.ps1 and read by
    Start-TablespacePreCheck.ps1 before the import begins.

    Called automatically at the end of Start-OracleExport.ps1, or run standalone.

.PARAMETER ConfigPath
    Path to ExportConfig.json. Defaults to .\Config\ExportConfig.json.

.EXAMPLE
    .\New-TablespaceSnapshot.ps1
    .\New-TablespaceSnapshot.ps1 -ConfigPath "C:\Scripts\Export\Config\ExportConfig.json"

.NOTES
    Compatible : PowerShell 5.0+
    Runs on    : aazeud-oracle02 (Production)
    Author     : DBA Team - Ashley Furniture India
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = "$PSScriptRoot\Config\ExportConfig.json"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Write-Info { param([string]$Msg) Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [TSSNAPSHOT] $Msg" -ForegroundColor Cyan }
function Write-Fail { param([string]$Msg) Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [TSSNAPSHOT] $Msg" -ForegroundColor Red }

try {
    if (-not (Test-Path $ConfigPath)) { throw "Config not found: $ConfigPath" }
    $config     = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
    $sqlplusExe = Join-Path $config.OracleHome 'bin\sqlplus.exe'
    $outPath    = Join-Path $config.ExportBaseDirectory 'TablespaceSnapshot.json'

    Write-Info "Capturing Production tablespace snapshot..."

    $sql = @"
SET PAGESIZE 0
SET LINESIZE 2000
SET FEEDBACK OFF
SET HEADING OFF
SET VERIFY OFF
SET TRIMOUT ON
WHENEVER SQLERROR EXIT SQL.SQLCODE

-- Tablespace summary
SELECT 'TS|' || t.tablespace_name || '|' ||
       TO_CHAR(ROUND(t.allocated_mb,1)) || '|' ||
       TO_CHAR(ROUND(t.used_mb,1))      || '|' ||
       TO_CHAR(ROUND(t.free_mb,1))      || '|' ||
       TO_CHAR(ROUND(t.pct_used,1))     || '|' ||
       t.autoextensible                 || '|' ||
       TO_CHAR(ROUND(t.max_mb,1))
FROM (
    SELECT df.tablespace_name,
           df.alloc_mb                                    AS allocated_mb,
           df.alloc_mb - NVL(fs.free_mb,0)               AS used_mb,
           NVL(fs.free_mb,0)                              AS free_mb,
           (df.alloc_mb - NVL(fs.free_mb,0))
             / NULLIF(df.alloc_mb,0) * 100                AS pct_used,
           df.autoextensible,
           df.max_mb
    FROM  (SELECT tablespace_name,
                  SUM(bytes)/1048576                           AS alloc_mb,
                  MAX(autoextensible)                          AS autoextensible,
                  SUM(DECODE(maxbytes,0,bytes,maxbytes))/1048576 AS max_mb
           FROM   dba_data_files
           GROUP  BY tablespace_name) df
    LEFT  JOIN
          (SELECT tablespace_name, SUM(bytes)/1048576 AS free_mb
           FROM   dba_free_space
           GROUP  BY tablespace_name) fs
    ON    df.tablespace_name = fs.tablespace_name
) t
ORDER BY t.tablespace_name;

-- Datafile detail
SELECT 'DF|' || tablespace_name || '|' || file_name || '|' ||
       TO_CHAR(ROUND(bytes/1048576,1)) || '|' ||
       autoextensible || '|' ||
       TO_CHAR(ROUND(DECODE(maxbytes,0,bytes,maxbytes)/1048576,1))
FROM   dba_data_files
ORDER  BY tablespace_name, file_id;

EXIT;
"@

    $tmpSql = Join-Path $env:TEMP "ts_snapshot_$(Get-Date -Format 'yyyyMMddHHmmss').sql"
    $tmpOut = "$tmpSql.out"
    try {
        $sql | Set-Content -Path $tmpSql -Encoding ASCII
        $proc = Start-Process -FilePath $sqlplusExe `
                              -ArgumentList "-S `"/ as sysdba`" @`"$tmpSql`"" `
                              -NoNewWindow -Wait -PassThru `
                              -RedirectStandardOutput $tmpOut
        if ($proc.ExitCode -ne 0) { throw "sqlplus exited with code $($proc.ExitCode)" }
        $lines = Get-Content $tmpOut | Where-Object { $_.Trim() -ne '' }
    }
    finally { Remove-Item $tmpSql,$tmpOut -Force -ErrorAction SilentlyContinue }

    # Parse output into structured objects
    $tablespaces = @{}
    $datafiles   = @{}

    foreach ($line in $lines) {
        $parts = $line.Trim() -split '\|'
        if ($parts[0] -eq 'TS' -and $parts.Count -ge 8) {
            $name = $parts[1]
            $tablespaces[$name] = @{
                Name        = $name
                AllocatedMB = [double]$parts[2]
                UsedMB      = [double]$parts[3]
                FreeMB      = [double]$parts[4]
                PctUsed     = [double]$parts[5]
                AutoExtend  = $parts[6]
                MaxMB       = [double]$parts[7]
            }
        }
        elseif ($parts[0] -eq 'DF' -and $parts.Count -ge 6) {
            $tsName = $parts[1]
            if (-not $datafiles.ContainsKey($tsName)) { $datafiles[$tsName] = @() }
            $datafiles[$tsName] += @{
                FileName    = $parts[2]
                SizeMB      = [double]$parts[3]
                AutoExtend  = $parts[4]
                MaxMB       = [double]$parts[5]
            }
        }
    }

    $snapshot = @{
        CapturedAt   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Server       = $env:COMPUTERNAME
        Tablespaces  = $tablespaces
        Datafiles    = $datafiles
    }

    $snapshot | ConvertTo-Json -Depth 6 | Set-Content -Path $outPath -Encoding UTF8
    Write-Info "Snapshot saved: $outPath ($($tablespaces.Count) tablespaces, $($datafiles.Values | ForEach-Object { $_.Count } | Measure-Object -Sum | Select-Object -ExpandProperty Sum) datafiles)"
}
catch {
    Write-Fail "Snapshot failed: $_"
    exit 1
}
