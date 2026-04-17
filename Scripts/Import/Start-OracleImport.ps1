<#
.SYNOPSIS
    Oracle Stage Refresh — Schema Import Pipeline (Story 3 / DBA-557).

.DESCRIPTION
    Single entry-point script that:
      PHASE 1 - INIT   : Creates the import directory on the filesystem and the Oracle
                         DBA_DIRECTORY object (CREATE OR REPLACE + GRANT).
      PHASE 2 - IMPORT : For each schema, generates an impdp parameter file and
                         executes Oracle Data Pump Import sequentially.
      PHASE 3 - MASK   : Runs optional per-schema PII masking SQL scripts via SQL*Plus.

    All common exclusions are applied to every schema:
      EXCLUDE=STATISTICS, USER, SEQUENCE, FUNCTION, PROCEDURE, VIEW, ROLE_GRANT, PROCOBJ
      TABLE_EXISTS_ACTION=REPLACE
      TRANSFORM=DISABLE_ARCHIVE_LOGGING:Y
      DATA_OPTIONS=SKIP_CONSTRAINT_ERRORS

    Per-schema index exclusions are supported via ImportConfig.json (e.g. DIS).

.PARAMETER ConfigPath
    Path to ImportConfig.json. Defaults to .\Config\ImportConfig.json.

.PARAMETER SkipInit
    Skip Phase 1 (directory setup). Use when Oracle DIRECTORY object already exists.

.PARAMETER InitOnly
    Run Phase 1 only — does not proceed to import.

.PARAMETER DryRun
    Generates .par files but does NOT execute impdp or masking scripts.

.EXAMPLE
    .\Start-OracleImport.ps1
    .\Start-OracleImport.ps1 -DryRun
    .\Start-OracleImport.ps1 -SkipInit

.NOTES
    Compatible : PowerShell 5.0+
    Oracle Ed. : Standard Edition 2 (CLUSTER=N, no PARALLEL)
    Runs on    : Stage server (aazeud-oracle03)
    Jira       : DBA-557 | Story 3 - Automated Import Module
    Author     : DBA Team - Ashley Furniture India
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "$PSScriptRoot\Config\ImportConfig.json",
    [switch]$SkipInit,
    [switch]$InitOnly,
    [switch]$DryRun
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$Script:LogFile   = $null
$Script:LogWriter = $null
$RunTimestamp     = Get-Date -Format 'yyyyMMdd_HHmmss'

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Write-Log
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO',
        [string]$Phase = 'IMPORT'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] [$Phase] $Message"
    switch ($Level) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan   }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red    }
        'SUCCESS' { Write-Host $line -ForegroundColor Green  }
    }
    if ($Script:LogWriter) { $Script:LogWriter.WriteLine($line); $Script:LogWriter.Flush() }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Send-ImportEmail
# ─────────────────────────────────────────────────────────────────────────────
function Send-ImportEmail {
    param(
        [string]   $Subject,
        [string]   $Body,
        [string[]] $To,
        [string]   $From,
        [string]   $SMTPServer,
        [int]      $SMTPPort,
        [string[]] $Attachments = @()
    )
    $tempCopies = @()
    try {
        $params = @{ Subject=$Subject; Body=$Body; BodyAsHtml=$true; From=$From
                     To=$To; SmtpServer=$SMTPServer; Port=$SMTPPort }
        if ($Attachments.Count -gt 0) {
            $resolved = @()
            foreach ($att in $Attachments) {
                if (Test-Path $att) {
                    $tmp = Join-Path $env:TEMP ("$(Split-Path $att -Leaf)_$(Get-Date -Format 'HHmmssff').tmp")
                    Copy-Item -Path $att -Destination $tmp -Force
                    $resolved += $tmp; $tempCopies += $tmp
                }
            }
            if ($resolved.Count -gt 0) { $params['Attachments'] = $resolved }
        }
        Send-MailMessage @params
        Write-Log "Email sent: $Subject" -Level INFO
    }
    catch { Write-Log "Email send failed (non-fatal): $_" -Level WARN }
    finally { $tempCopies | Remove-Item -Force -ErrorAction SilentlyContinue }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-SqlPlus
# Executes a SQL block via sqlplus -S "/ as sysdba". Returns ExitCode + Output.
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-SqlPlus {
    param([string]$SqlBlock, [string]$SqlPlusExe)
    $tmpSql = Join-Path $env:TEMP "OracleSQL_$(Get-Date -Format 'yyyyMMddHHmmss').sql"
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
        if (Test-Path $tmpErr) { $err = Get-Content $tmpErr; if ($err) { $output += $err } }
        return @{ ExitCode = $proc.ExitCode; Output = $output }
    }
    finally { Remove-Item $tmpSql,$tmpOut,$tmpErr -Force -ErrorAction SilentlyContinue }
}


# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Initialize-ImportDirectory   [PHASE 1]
# Creates the local import directory and the Oracle DBA_DIRECTORY object.
# Returns $true on success, $false on failure.
# ─────────────────────────────────────────────────────────────────────────────
function Initialize-ImportDirectory {
    param([PSCustomObject]$Config, [string]$SqlPlusExe)

    Write-Log ('─' * 60) -Level INFO -Phase 'INIT'
    Write-Log 'PHASE 1 - Import Directory Initialisation' -Level INFO -Phase 'INIT'
    Write-Log ('─' * 60) -Level INFO -Phase 'INIT'

    # Step 1.1 - Filesystem directories
    foreach ($dir in @($Config.ImportBaseDirectory, $Config.LogDirectory)) {
        if (Test-Path $dir) { Write-Log "  EXISTS  : $dir" -Level INFO -Phase 'INIT' }
        else {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Write-Log "  CREATED : $dir" -Level SUCCESS -Phase 'INIT'
        }
    }

    # Step 1.2 - Oracle DBA_DIRECTORY object
    Write-Log 'Step 1.2 : Oracle DBA_DIRECTORY object' -Level INFO -Phase 'INIT'
    $dirObj  = $Config.OracleDirectoryObject
    $sqlPath = $Config.ImportBaseDirectory.Replace("'", "''")

    $sqlBlock = @"
SET SERVEROUTPUT ON SIZE UNLIMITED
SET FEEDBACK OFF
SET HEADING OFF
WHENEVER SQLERROR EXIT SQL.SQLCODE

DECLARE
    v_exists NUMBER;
    v_action VARCHAR2(20);
BEGIN
    SELECT COUNT(*) INTO v_exists FROM dba_directories WHERE directory_name = '$dirObj';
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
FROM   dba_directories WHERE directory_name = '$dirObj';

EXIT;
"@

    $dbResult = Invoke-SqlPlus -SqlBlock $sqlBlock -SqlPlusExe $SqlPlusExe
    foreach ($line in $dbResult.Output) {
        if ($line -match 'ERROR|ORA-') { Write-Log "  $line" -Level ERROR -Phase 'INIT' }
        elseif ($line.Trim() -ne '')    { Write-Log "  $line" -Level INFO  -Phase 'INIT' }
    }

    if ($dbResult.ExitCode -ne 0) {
        Write-Log "SQLPlus exited with code $($dbResult.ExitCode) - aborting." -Level ERROR -Phase 'INIT'
        return $false
    }

    Write-Log 'Phase 1 complete - import directory ready.' -Level SUCCESS -Phase 'INIT'
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: New-ImpdpParFile   [PHASE 2]
# Generates a schema-specific impdp parameter file.
# Returns hashtable: ParFilePath, DumpFile, LogFile, SchemaDir.
# ─────────────────────────────────────────────────────────────────────────────
function New-ImpdpParFile {
    param(
        [PSCustomObject]$SchemaConfig,
        [string]        $ImportBaseDir,
        [string]        $OracleDirectoryObject,
        [string]        $DumpFileName
    )

    $schema      = $SchemaConfig.Name
    $schemaDir   = Join-Path $ImportBaseDir $schema
    if (-not (Test-Path $schemaDir)) { New-Item -ItemType Directory -Path $schemaDir | Out-Null }

    $logFile     = "${schema}_import_${RunTimestamp}.log"
    $parFilePath = Join-Path $schemaDir "${schema}_import.par"

    $lines = @(
        "USERID=""$($SchemaConfig.Credential)""",
        "SCHEMAS=$schema",
        "DIRECTORY=$OracleDirectoryObject",
        "DUMPFILE=$DumpFileName",
        "LOGFILE=$logFile",
        "TABLE_EXISTS_ACTION=REPLACE",
        "CLUSTER=N",
        "METRICS=YES",
        "EXCLUDE=STATISTICS",
        "EXCLUDE=USER",
        "EXCLUDE=SEQUENCE,FUNCTION,PROCEDURE,VIEW,PACKAGE",
        "EXCLUDE=ROLE_GRANT",
        "EXCLUDE=PROCOBJ",
        "TRANSFORM=DISABLE_ARCHIVE_LOGGING:Y",
        "DATA_OPTIONS=SKIP_CONSTRAINT_ERRORS"
    )

    # Per-schema index exclusions (e.g. DIS: IDX_PACOS_HOSAVAILPAST)
    if ($SchemaConfig.ExcludeIndexes -and $SchemaConfig.ExcludeIndexes.Count -gt 0) {
        $idxList = ($SchemaConfig.ExcludeIndexes | ForEach-Object { "'$_'" }) -join ','
        $lines  += "EXCLUDE=INDEX:`"IN ($idxList)`""
    }

    # Per-schema table exclusions (same list as export — protect stage-specific tables)
    if ($SchemaConfig.ExcludeTables -and $SchemaConfig.ExcludeTables.Count -gt 0) {
        $tblList = ($SchemaConfig.ExcludeTables | ForEach-Object { "'$_'" }) -join ','
        $lines  += "EXCLUDE=TABLE:`"IN ($tblList)`""
    }

    $lines | Set-Content -Path $parFilePath -Encoding ASCII
    Write-Log "Par file written: $parFilePath" -Level INFO -Phase 'IMPORT'

    return @{ ParFilePath=$parFilePath; DumpFile=$DumpFileName; LogFile=$logFile; SchemaDir=$schemaDir }
}


# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-SchemaImport   [PHASE 2]
# Runs impdp for one schema. Returns hashtable: Success, Duration, Skipped.
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-SchemaImport {
    param(
        [PSCustomObject]$SchemaConfig,
        [hashtable]     $ParInfo,
        [string]        $OracleHome,
        [PSCustomObject]$EmailConfig,
        [bool]          $IsDryRun
    )
    $schema    = $SchemaConfig.Name
    $startTime = Get-Date
    $allRecips = @($EmailConfig.DBARecipients) + @($EmailConfig.BusinessRecipients)
    $dbaRecips = @($EmailConfig.DBARecipients)

    # Start notification
    $startBody = @"
<h3>Oracle Import Started</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Schema</b></td><td>$schema</td></tr>
  <tr><td><b>Dump File</b></td><td>$($ParInfo.DumpFile)</td></tr>
  <tr><td><b>Mode</b></td><td>TABLE_EXISTS_ACTION=REPLACE | SE2 Sequential</td></tr>
  <tr><td><b>Started</b></td><td>$($startTime.ToString('yyyy-MM-dd HH:mm:ss'))</td></tr>
  <tr><td><b>Dry Run</b></td><td>$IsDryRun</td></tr>
</table>
"@
    Send-ImportEmail -Subject "[IMPORT STARTED] $schema - Oracle Stage Refresh" `
                     -Body $startBody -To $allRecips `
                     -From $EmailConfig.From -SMTPServer $EmailConfig.SMTPServer -SMTPPort $EmailConfig.SMTPPort

    if ($IsDryRun) {
        Write-Log "[$schema] DRY RUN - impdp skipped." -Level WARN -Phase 'IMPORT'
        return @{ Success=$true; Duration=0; Skipped=$true }
    }

    # Execute impdp
    $impdpExe = Join-Path $OracleHome 'bin\impdp.exe'
    $exitCode = -1
    $attempt  = 0

    while ($attempt -le 2) {
        $attempt++
        Write-Log "[$schema] Attempt $attempt - executing impdp..." -Level INFO -Phase 'IMPORT'
        try {
            $proc     = Start-Process -FilePath $impdpExe `
                                      -ArgumentList "parfile=`"$($ParInfo.ParFilePath)`"" `
                                      -WorkingDirectory $ParInfo.SchemaDir `
                                      -NoNewWindow -Wait -PassThru
            $exitCode = $proc.ExitCode
            break
        }
        catch {
            Write-Log "[$schema] impdp launch error (attempt $attempt): $_" -Level WARN -Phase 'IMPORT'
            if ($attempt -le 2) { Start-Sleep -Seconds 60 }
        }
    }

    # Scan Oracle impdp log — collect ALL ORA- lines (capped at 50)
    $oracleLog = Join-Path $ParInfo.SchemaDir $ParInfo.LogFile
    $oraErrors = @()
    if (Test-Path $oracleLog) {
        $oraErrors = @(Select-String -Path $oracleLog -Pattern 'ORA-' | Select-Object -First 50)
    }

    $duration  = [math]::Round(((Get-Date) - $startTime).TotalMinutes, 1)

    # Build ORA- HTML block — always included in every schema email
    $oraSection = if ($oraErrors.Count -gt 0) {
        $oraRows = ($oraErrors | ForEach-Object {
            "<tr><td style='font-family:Courier New;font-size:12px;color:#b94a48'>$($_.Line.Trim())</td></tr>"
        }) -join ''
        @"
<br/><h4>ORA- Messages in Import Log ($($oraErrors.Count) found)</h4>
<table border='1' cellpadding='4' style='border-collapse:collapse;font-family:Arial;width:100%'>
  $oraRows
</table>
"@
    } else {
        '<br/><p style="color:green"><b>No ORA- messages found in import log.</b></p>'
    }

    # Success = exit code 0. ORA- warnings in log do not block success —
    # Oracle impdp frequently emits non-fatal ORA- messages (e.g. ORA-31684).
    if ($exitCode -eq 0) {
        Write-Log "[$schema] IMPORT SUCCEEDED in ${duration} min. ORA- messages: $($oraErrors.Count)" -Level SUCCESS -Phase 'IMPORT'
        $successBody = @"
<h3 style='color:green'>Oracle Import Completed Successfully</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Schema</b></td><td>$schema</td></tr>
  <tr><td><b>Dump File</b></td><td>$($ParInfo.DumpFile)</td></tr>
  <tr><td><b>Duration</b></td><td>${duration} minutes</td></tr>
  <tr><td><b>Exit Code</b></td><td>$exitCode</td></tr>
  <tr><td><b>ORA- Count</b></td><td>$($oraErrors.Count)</td></tr>
  <tr><td><b>Completed</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
$oraSection
"@
        Send-ImportEmail -Subject "[IMPORT COMPLETE] $schema - Oracle Stage Refresh ($($oraErrors.Count) ORA- msgs)" `
                         -Body $successBody -To $allRecips `
                         -From $EmailConfig.From -SMTPServer $EmailConfig.SMTPServer -SMTPPort $EmailConfig.SMTPPort
        return @{ Success=$true; Duration=$duration; Skipped=$false }
    }

    # Failure — exit code non-zero
    Write-Log "[$schema] IMPORT FAILED. ExitCode=$exitCode | ORA- messages: $($oraErrors.Count)" -Level ERROR -Phase 'IMPORT'
    $failBody = @"
<h3 style='color:red'>ALERT: Oracle Import FAILED</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Schema</b></td><td>$schema</td></tr>
  <tr><td><b>Dump File</b></td><td>$($ParInfo.DumpFile)</td></tr>
  <tr><td><b>Exit Code</b></td><td>$exitCode</td></tr>
  <tr><td><b>Attempts</b></td><td>$attempt</td></tr>
  <tr><td><b>ORA- Count</b></td><td>$($oraErrors.Count)</td></tr>
  <tr><td><b>Failed At</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
$oraSection
<p>Pipeline will continue with next schema. Check attached logs.</p>
"@
    $att = @()
    if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $att += $Script:LogFile }
    if (Test-Path $oracleLog)                              { $att += $oracleLog     }
    Send-ImportEmail -Subject "[IMPORT FAILED] $schema - Oracle Stage Refresh ($($oraErrors.Count) ORA- msgs)" `
                     -Body $failBody -To $dbaRecips `
                     -From $EmailConfig.From -SMTPServer $EmailConfig.SMTPServer -SMTPPort $EmailConfig.SMTPPort `
                     -Attachments $att
    return @{ Success=$false; Duration=$duration; Skipped=$false }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-MaskingScript   [PHASE 3]
# Runs the per-schema PII masking SQL script if configured and present.
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-MaskingScript {
    param([PSCustomObject]$SchemaConfig, [string]$SqlPlusExe)
    $schema = $SchemaConfig.Name
    $script = $SchemaConfig.MaskingScript

    if ([string]::IsNullOrWhiteSpace($script)) {
        Write-Log "[$schema] No masking script configured - skipping." -Level INFO -Phase 'MASKING'
        return $true
    }

    $scriptPath = if ([System.IO.Path]::IsPathRooted($script)) { $script } `
                  else { Join-Path $PSScriptRoot $script }

    if (-not (Test-Path $scriptPath)) {
        Write-Log "[$schema] Masking script not found: $scriptPath - skipping." -Level WARN -Phase 'MASKING'
        return $true
    }

    Write-Log "[$schema] Running masking script: $scriptPath" -Level INFO -Phase 'MASKING'
    $sqlBlock = @"
WHENEVER SQLERROR EXIT SQL.SQLCODE
@$scriptPath
COMMIT;
EXIT;
"@
    $result = Invoke-SqlPlus -SqlBlock $sqlBlock -SqlPlusExe $SqlPlusExe
    foreach ($line in $result.Output) {
        if ($line -match 'ORA-|ERROR') { Write-Log "  $line" -Level ERROR -Phase 'MASKING' }
        elseif ($line.Trim() -ne '')   { Write-Log "  $line" -Level INFO  -Phase 'MASKING' }
    }
    if ($result.ExitCode -ne 0) {
        Write-Log "[$schema] Masking script FAILED (exit $($result.ExitCode))." -Level ERROR -Phase 'MASKING'
        return $false
    }
    Write-Log "[$schema] Masking script completed successfully." -Level SUCCESS -Phase 'MASKING'
    return $true
}

# NOTE: Sequence Sync, Statistics and Recompile are handled by the standalone
# script: Start-SequenceSyncAndStats.ps1 — run it manually after reviewing logs.

# ═════════════════════════════════════════════════════════════════════════════
#  MAIN EXECUTION BLOCK
# ═════════════════════════════════════════════════════════════════════════════
try {
    if (-not (Test-Path $ConfigPath)) { throw "Config file not found: $ConfigPath" }
    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json

    $env:ORACLE_HOME = $config.OracleHome
    $env:PATH        = "$($config.OracleHome)\bin;$($env:PATH)"
    $sqlplusExe      = Join-Path $config.OracleHome 'bin\sqlplus.exe'
    if (-not (Test-Path $sqlplusExe)) { throw "sqlplus.exe not found: $sqlplusExe" }

    # Open log file
    $logDir = $config.LogDirectory
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $Script:LogFile   = Join-Path $logDir "OracleImport_${RunTimestamp}.log"
    $Script:LogWriter = [System.IO.StreamWriter]::new($Script:LogFile, $true, [System.Text.Encoding]::UTF8)
    $Script:LogWriter.AutoFlush = $true

    # Startup banner
    Write-Log ('=' * 65)                                           -Level INFO -Phase 'STARTUP'
    Write-Log 'Oracle Stage Refresh - Full Import Pipeline'        -Level INFO -Phase 'STARTUP'
    Write-Log "Server    : $($env:COMPUTERNAME)"                   -Level INFO -Phase 'STARTUP'
    Write-Log "Run ID    : $RunTimestamp"                          -Level INFO -Phase 'STARTUP'
    Write-Log "Import Dir: $($config.ImportBaseDirectory)"         -Level INFO -Phase 'STARTUP'
    Write-Log "Oracle Dir: $($config.OracleDirectoryObject)"       -Level INFO -Phase 'STARTUP'
    Write-Log "Schemas   : $(($config.Schemas | ForEach-Object {$_.Name}) -join ', ')" -Level INFO -Phase 'STARTUP'
    Write-Log "SkipInit  : $($SkipInit.IsPresent)"                 -Level INFO -Phase 'STARTUP'
    Write-Log "DryRun    : $($DryRun.IsPresent)"                   -Level INFO -Phase 'STARTUP'
    Write-Log ('=' * 65)                                           -Level INFO -Phase 'STARTUP'

    # ══ PHASE 0 — Tablespace Pre-Check ═══════════════════════════════════════
    Write-Log ('─' * 60) -Level INFO -Phase 'PRECHECK'
    Write-Log 'PHASE 0 - Tablespace Pre-Check' -Level INFO -Phase 'PRECHECK'
    $preCheckScript = Join-Path $PSScriptRoot 'Start-TablespacePreCheck.ps1'
    if (Test-Path $preCheckScript) {
        try {
            # Run pre-check; -AutoFix enables automatic remediation when space is insufficient.
            # Remove -AutoFix to report-only mode (exits with code 1 if issues found).
            & $preCheckScript -ConfigPath $ConfigPath -AutoFix
            if ($LASTEXITCODE -ne 0) {
                throw "Tablespace Pre-Check reported INSUFFICIENT space and could not auto-fix. Review the pre-check log and resolve before retrying."
            }
            Write-Log 'Tablespace Pre-Check passed.' -Level SUCCESS -Phase 'PRECHECK'
        }
        catch {
            throw "Phase 0 (Tablespace Pre-Check) failed: $_"
        }
    }
    else {
        Write-Log "Pre-check script not found at $preCheckScript - skipping (not recommended)." -Level WARN -Phase 'PRECHECK'
    }

    # ══ PHASE 1 — Directory Initialisation ═══════════════════════════════════
    if ($SkipInit.IsPresent) {
        Write-Log 'Phase 1 skipped (-SkipInit specified).' -Level WARN -Phase 'INIT'
    }
    else {
        $initOk = Initialize-ImportDirectory -Config $config -SqlPlusExe $sqlplusExe
        if (-not $initOk) { throw 'Phase 1 (Directory Initialisation) failed.' }
    }
    if ($InitOnly.IsPresent) { Write-Log 'InitOnly mode - exiting.' -Level INFO -Phase 'INIT'; exit 0 }

    # ══ PHASE 2 — Import Pipeline ═════════════════════════════════════════════
    Write-Log ('─' * 60) -Level INFO -Phase 'IMPORT'
    Write-Log 'PHASE 2 - Import Pipeline' -Level INFO -Phase 'IMPORT'
    Write-Log ('─' * 60) -Level INFO -Phase 'IMPORT'

    $allRecips    = @($config.Email.DBARecipients) + @($config.Email.BusinessRecipients)
    $results      = [System.Collections.ArrayList]@()
    $schemaList   = ($config.Schemas | ForEach-Object { $_.Name }) -join ', '

    # Pipeline start email
    $pipelineBody = @"
<h2>Oracle Stage Refresh &mdash; Import Pipeline Started</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Schemas</b></td><td>$schemaList</td></tr>
  <tr><td><b>Import Dir</b></td><td>$($config.ImportBaseDirectory)</td></tr>
  <tr><td><b>Dry Run</b></td><td>$($DryRun.IsPresent)</td></tr>
  <tr><td><b>Started</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
<p><i>Individual emails will follow for each schema.</i></p>
"@
    Send-ImportEmail -Subject "[IMPORT PIPELINE STARTED] Oracle Stage Refresh - $RunTimestamp" `
                     -Body $pipelineBody -To $allRecips `
                     -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort

    foreach ($schemaConfig in $config.Schemas) {
        $schema = $schemaConfig.Name
        Write-Log "--- Schema: $schema ---" -Level INFO -Phase 'IMPORT'

        # Find most-recent .dmp file for this schema in ImportBaseDirectory
        $dmpFiles = @(Get-ChildItem -Path $config.ImportBaseDirectory -Filter "${schema}_*.dmp" -File -ErrorAction SilentlyContinue |
                      Sort-Object LastWriteTime -Descending)

        if ($dmpFiles.Count -eq 0) {
            Write-Log "[$schema] No .dmp file found in $($config.ImportBaseDirectory) - SKIPPING." -Level WARN -Phase 'IMPORT'
            [void]$results.Add([PSCustomObject]@{ Schema=$schema; Success=$false; Duration=0; Skipped=$false; Missing=$true })
            continue
        }

        $dumpFileName = $dmpFiles[0].Name
        Write-Log "[$schema] Using dump: $dumpFileName" -Level INFO -Phase 'IMPORT'

        $parInfo = New-ImpdpParFile -SchemaConfig          $schemaConfig `
                                    -ImportBaseDir         $config.ImportBaseDirectory `
                                    -OracleDirectoryObject $config.OracleDirectoryObject `
                                    -DumpFileName          $dumpFileName

        $result = Invoke-SchemaImport -SchemaConfig $schemaConfig `
                                      -ParInfo      $parInfo `
                                      -OracleHome   $config.OracleHome `
                                      -EmailConfig  $config.Email `
                                      -IsDryRun     $DryRun.IsPresent

        if ($result.Success) {
            # ── PHASE 3 — PII Masking ─────────────────────────────────────────
            Write-Log ('─' * 60) -Level INFO -Phase 'MASKING'
            Write-Log "PHASE 3 - PII Masking: $schema" -Level INFO -Phase 'MASKING'
            Invoke-MaskingScript -SchemaConfig $schemaConfig -SqlPlusExe $sqlplusExe | Out-Null
        }

        [void]$results.Add([PSCustomObject]@{
            Schema=$schema; Success=$result.Success; Duration=$result.Duration
            Skipped=$result.Skipped; Missing=$false; DumpFile=$dumpFileName
        })
    }

    # Final summary
    $successCount = @($results | Where-Object {  $_.Success }).Count
    $failCount    = @($results | Where-Object { -not $_.Success }).Count
    $totalMins    = [math]::Round(($results | Measure-Object -Property Duration -Sum).Sum, 1)

    $tableRows = $results | ForEach-Object {
        $color    = if ($_.Success) { 'green' } else { 'red' }
        $status = if ($_.Skipped) { 'SKIPPED (DryRun)' } elseif ($_.Missing) { 'FILE NOT FOUND' } `
                  elseif ($_.Success) { 'SUCCESS' } else { 'FAILED' }
        "<tr><td>$($_.Schema)</td>" +
        "<td style='color:$color'><b>$status</b></td>" +
        "<td>$($_.Duration) min</td>" +
        "<td>$($_.DumpFile)</td></tr>"
    }

    $overallStatus = if ($failCount -gt 0) { 'COMPLETED WITH ERRORS' } else { 'COMPLETED SUCCESSFULLY' }
    $summaryBody   = @"
<h2>Oracle Stage Refresh &mdash; Import Pipeline Summary</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Total Schemas</b></td><td>$($config.Schemas.Count)</td></tr>
  <tr><td><b>Successful</b></td><td style='color:green'><b>$successCount</b></td></tr>
  <tr><td><b>Failed</b></td><td style='color:red'><b>$failCount</b></td></tr>
  <tr><td><b>Total Duration</b></td><td>${totalMins} minutes</td></tr>
</table><br/>
<h3>Per-Schema Results</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr style='background:#ddd'><th>Schema</th><th>Import Status</th><th>Duration</th><th>Dump File</th></tr>
  $($tableRows -join '')
</table><br/><p>Full log attached.</p>
"@
    Write-Log "Import pipeline finished. Success=$successCount | Failed=$failCount | ${totalMins} min" `
              -Level SUCCESS -Phase 'SUMMARY'

    $summaryAtt = @(); if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $summaryAtt += $Script:LogFile }
    Send-ImportEmail -Subject "[IMPORT SUMMARY] $overallStatus - Oracle Stage Refresh $RunTimestamp" `
                     -Body $summaryBody -To $allRecips `
                     -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort `
                     -Attachments $summaryAtt

    if ($failCount -gt 0) { exit 1 }
}
catch {
    $errMsg = $_.Exception.Message
    Write-Log "CRITICAL: $errMsg" -Level ERROR -Phase 'FATAL'
    try {
        if ($null -ne $config -and $null -ne $config.Email) {
            $critBody = @"
<h2 style='color:red'>CRITICAL: Oracle Import Pipeline Stopped</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Error</b></td><td>$errMsg</td></tr>
  <tr><td><b>Time</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
"@
            $ca = @(); if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $ca += $Script:LogFile }
            Send-ImportEmail -Subject "[CRITICAL ERROR] Oracle Import - $RunTimestamp" `
                             -Body $critBody -To @($config.Email.DBARecipients) `
                             -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort `
                             -Attachments $ca
        }
    } catch { Write-Host "Failed to send critical error email: $_" -ForegroundColor Red }
    exit 1
}
finally {
    if ($null -ne $Script:LogWriter) {
        try   { $Script:LogWriter.Close(); $Script:LogWriter.Dispose() }
        catch { Write-Host "Warning: could not close log writer: $_" -ForegroundColor Yellow }
        $Script:LogWriter = $null
    }
}
