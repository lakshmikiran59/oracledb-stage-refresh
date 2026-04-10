<#
.SYNOPSIS
    Creates physical export directories on the server and the Oracle DBA_DIRECTORY object in the database.

.DESCRIPTION
    Reads ExportConfig.json and performs two operations:
      1. FILESYSTEM  - Creates the base export dir, log dir, and a subfolder per schema if they do not exist.
      2. DATABASE    - Creates (or replaces) the Oracle DIRECTORY object via SQLPlus as SYSDBA,
                       then grants READ and WRITE on it to PUBLIC, and verifies the result.

.PARAMETER ConfigPath
    Path to ExportConfig.json. Defaults to .\Config\ExportConfig.json.

.PARAMETER WhatIf
    Shows what would be created without actually creating anything (filesystem or DB).

.EXAMPLE
    .\Initialize-ExportDirectories.ps1
    .\Initialize-ExportDirectories.ps1 -WhatIf
    .\Initialize-ExportDirectories.ps1 -ConfigPath "D:\Custom\ExportConfig.json"

.NOTES
    Compatible : PowerShell 5.0+
    Runs as    : Must be executed by an OS account with write access to the export drive
                 AND Oracle SYSDBA privilege (/ as sysdba).
    Jira       : DBA-555 | Sub-task 5.1 - Pre-Refresh Environment Setup
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "$PSScriptRoot\Config\ExportConfig.json",
    [switch]$WhatIf
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Write-Step
# Consistent colour-coded console output for each initialisation step.
# ─────────────────────────────────────────────────────────────────────────────
function Write-Step {
    param(
        [string]$Message,
        [ValidateSet('Cyan','Green','Yellow','Red','White')][string]$Color = 'White'
    )
    $ts = Get-Date -Format 'HH:mm:ss'
    Write-Host "[$ts] $Message" -ForegroundColor $Color
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-SqlPlus
# Writes a SQL block to a temp file and executes it via sqlplus -S / as sysdba.
# Returns the stdout output as a string array and sets $LastExitCode.
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-SqlPlus {
    param(
        [string]$SqlBlock,
        [string]$SqlPlusExe
    )
    $tmpSql = Join-Path $env:TEMP "OracleInit_$(Get-Date -Format 'yyyyMMddHHmmss').sql"
    $tmpOut = "$tmpSql.out"
    $tmpErr = "$tmpSql.err"

    try {
        $SqlBlock | Set-Content -Path $tmpSql -Encoding ASCII

        $proc = Start-Process -FilePath $SqlPlusExe `
                              -ArgumentList "-S `"/ as sysdba`" @`"$tmpSql`"" `
                              -NoNewWindow -Wait -PassThru `
                              -RedirectStandardOutput $tmpOut `
                              -RedirectStandardError  $tmpErr

        $output = @()
        if (Test-Path $tmpOut) { $output += Get-Content $tmpOut }
        if (Test-Path $tmpErr) {
            $errContent = Get-Content $tmpErr
            if ($errContent) { $output += $errContent }
        }

        return @{ ExitCode = $proc.ExitCode; Output = $output }
    }
    finally {
        Remove-Item $tmpSql, $tmpOut, $tmpErr -Force -ErrorAction SilentlyContinue
    }
}

# ═════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═════════════════════════════════════════════════════════════════════════════
Write-Step ('=' * 60) -Color Cyan
Write-Step 'Oracle Export - Directory Initialisation' -Color Cyan
Write-Step "Config : $ConfigPath" -Color Cyan
Write-Step "WhatIf : $($WhatIf.IsPresent)" -Color Cyan
Write-Step ('=' * 60) -Color Cyan

# ── Load configuration ────────────────────────────────────────────────────────
if (-not (Test-Path $ConfigPath)) { throw "Config not found: $ConfigPath" }
$config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json

$env:ORACLE_HOME = $config.OracleHome
$env:PATH        = "$($config.OracleHome)\bin;$env:PATH"
$sqlplusExe      = Join-Path $config.OracleHome 'bin\sqlplus.exe'

if (-not (Test-Path $sqlplusExe)) { throw "sqlplus.exe not found at: $sqlplusExe" }

# ── STEP 1: Physical filesystem directories ───────────────────────────────────
Write-Step '' -Color White
Write-Step '-- STEP 1: Filesystem Directories --' -Color Cyan

# Build the full list: base dir + logs dir + one subdir per schema
$dirsToCreate  = [System.Collections.ArrayList]@()
[void]$dirsToCreate.Add($config.ExportBaseDirectory)
[void]$dirsToCreate.Add($config.LogDirectory)
foreach ($schema in $config.Schemas) {
    [void]$dirsToCreate.Add((Join-Path $config.ExportBaseDirectory $schema.Name))
}

$fsResults = [System.Collections.ArrayList]@()
foreach ($dir in $dirsToCreate) {
    if (Test-Path $dir) {
        Write-Step "  EXISTS   : $dir" -Color Green
        [void]$fsResults.Add([PSCustomObject]@{ Path = $dir; Action = 'Already Existed' })
    }
    else {
        if ($WhatIf) {
            Write-Step "  WHATIF   : Would create -> $dir" -Color Yellow
            [void]$fsResults.Add([PSCustomObject]@{ Path = $dir; Action = 'Would Create (WhatIf)' })
        }
        else {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Write-Step "  CREATED  : $dir" -Color Green
            [void]$fsResults.Add([PSCustomObject]@{ Path = $dir; Action = 'Created' })
        }
    }
}


# ── STEP 2: Oracle DBA_DIRECTORY object ──────────────────────────────────────
Write-Step '' -Color White
Write-Step '-- STEP 2: Oracle DBA_DIRECTORY Object --' -Color Cyan

$dirObj  = $config.OracleDirectoryObject          # e.g. ORCL_DATAPUMP
$dirPath = $config.ExportBaseDirectory             # e.g. E:\Datapump
$sqlPath = $dirPath.Replace("'", "''")             # escape single-quotes for PL/SQL

$sqlBlock = @"
SET SERVEROUTPUT ON SIZE UNLIMITED
SET FEEDBACK OFF
SET HEADING OFF
WHENEVER SQLERROR EXIT SQL.SQLCODE

DECLARE
    v_exists  NUMBER;
    v_action  VARCHAR2(20);
BEGIN
    SELECT COUNT(*) INTO v_exists
    FROM   dba_directories
    WHERE  directory_name = '$dirObj';

    IF v_exists = 0 THEN
        EXECUTE IMMEDIATE 'CREATE DIRECTORY $dirObj AS ''$sqlPath''';
        v_action := 'CREATED';
    ELSE
        EXECUTE IMMEDIATE 'CREATE OR REPLACE DIRECTORY $dirObj AS ''$sqlPath''';
        v_action := 'REPLACED';
    END IF;

    DBMS_OUTPUT.PUT_LINE('DB_ACTION : ' || v_action || ' -> $dirObj');
END;
/

GRANT READ, WRITE ON DIRECTORY $dirObj TO PUBLIC;
PROMPT DB_GRANT  : READ, WRITE ON $dirObj granted to PUBLIC

SELECT 'DB_VERIFY : ' || directory_name || ' => ' || directory_path
FROM   dba_directories
WHERE  directory_name = '$dirObj';

EXIT;
"@

if ($WhatIf) {
    Write-Step "  WHATIF   : Would create/replace DIRECTORY $dirObj => $dirPath" -Color Yellow
    Write-Step "  WHATIF   : Would grant READ, WRITE ON DIRECTORY $dirObj TO PUBLIC" -Color Yellow
    $dbResult = @{ ExitCode = 0; Output = @('WhatIf - no SQL executed') }
}
else {
    Write-Step "  Connecting to Oracle (/ as sysdba)..." -Color White
    $dbResult = Invoke-SqlPlus -SqlBlock $sqlBlock -SqlPlusExe $sqlplusExe

    foreach ($line in $dbResult.Output) {
        if     ($line -match 'ERROR|ORA-') { Write-Step "  $line" -Color Red    }
        elseif ($line.Trim() -ne '')        { Write-Step "  $line" -Color Green  }
    }

    if ($dbResult.ExitCode -ne 0) {
        Write-Step "  SQLPlus exited with code $($dbResult.ExitCode). Review output above." -Color Red
        exit 1
    }
}

# ── SUMMARY ───────────────────────────────────────────────────────────────────
Write-Step '' -Color White
Write-Step ('=' * 60) -Color Cyan
Write-Step 'INITIALISATION SUMMARY' -Color Cyan
Write-Step ('=' * 60) -Color Cyan

Write-Step 'Filesystem Directories:' -Color White
foreach ($r in $fsResults) {
    Write-Step ("  {0,-45} [{1}]" -f $r.Path, $r.Action) -Color Green
}

Write-Step '' -Color White
Write-Step 'Oracle Directory Object:' -Color White
Write-Step ("  {0,-25} => {1}" -f $dirObj, $dirPath) -Color Green
Write-Step "  Permissions : READ, WRITE granted to PUBLIC" -Color Green

Write-Step '' -Color White
if ($WhatIf) {
    Write-Step 'WhatIf mode - no changes were made to filesystem or database.' -Color Yellow
}
else {
    Write-Step 'Initialisation complete. Environment is ready for the export.' -Color Green
    Write-Step 'Next step : .\Invoke-OracleExport.ps1 -DryRun' -Color Cyan
}
Write-Step ('=' * 60) -Color Cyan
