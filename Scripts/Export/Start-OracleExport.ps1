<#
.SYNOPSIS
    Combined Oracle Export Automation - Directory Initialisation + Schema Export Pipeline.

.DESCRIPTION
    Single entry-point script that:
      PHASE 1 - INIT   : Creates physical directories on the server and the Oracle
                         DBA_DIRECTORY object (CREATE OR REPLACE + GRANT).
      PHASE 2 - EXPORT : Generates schema-specific expdp parameter files and executes
                         Oracle Data Pump Export for all 8 schemas sequentially.
                         Sends HTML email notifications at every stage.

    Schemas : ASHLEY, CIM, DIS, EMANIFEST, FINANCE, LEAN, SQLORACLE, TAPLSQL
    Edition : Oracle Standard Edition 2 (CLUSTER=N, no PARALLEL, no COMPRESSION)

.PARAMETER ConfigPath
    Path to ExportConfig.json. Defaults to .\Config\ExportConfig.json.

.PARAMETER DryRun
    Phase 2 only - generates .par files but does NOT execute expdp.

.PARAMETER SkipInit
    Skip Phase 1 entirely. Use when directories and Oracle DIRECTORY object already exist.

.PARAMETER InitOnly
    Run Phase 1 only. Does not proceed to the export pipeline.

.EXAMPLE
    # Full run - init + export
    .\Start-OracleExport.ps1

    # First time setup preview
    .\Start-OracleExport.ps1 -InitOnly

    # Directories already exist - go straight to export
    .\Start-OracleExport.ps1 -SkipInit

    # Validate config and par files without running expdp
    .\Start-OracleExport.ps1 -DryRun

.NOTES
    Compatible : PowerShell 5.0+
    Jira       : DBA-555 | Story 1 - Automated Export Module
    Author     : DBA Team - Ashley Furniture India
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "$PSScriptRoot\Config\ExportConfig.json",
    [switch]$DryRun,
    [switch]$SkipInit,
    [switch]$InitOnly
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$Script:LogFile   = $null
$Script:LogWriter = $null
$RunTimestamp     = Get-Date -Format 'yyyyMMdd_HHmmss'

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Write-Log
# Unified timestamped colour-coded logger for both phases.
# Writes to console always; writes to log file once StreamWriter is open.
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO',
        [string]$Phase = 'EXPORT'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] [$Phase] $Message"

    switch ($Level) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan   }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red    }
        'SUCCESS' { Write-Host $line -ForegroundColor Green  }
    }
    if ($Script:LogWriter) {
        $Script:LogWriter.WriteLine($line)
        $Script:LogWriter.Flush()
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Send-ExportEmail
# Sends an HTML email via Send-MailMessage (PS 5.0 built-in).
# Copies attachments to temp files first to prevent file-handle locks when
# Send-MailMessage fails to dispose Attachment objects on SMTP errors.
# ─────────────────────────────────────────────────────────────────────────────
function Send-ExportEmail {
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
        $params = @{
            Subject    = $Subject
            Body       = $Body
            BodyAsHtml = $true
            From       = $From
            To         = $To
            SmtpServer = $SMTPServer
            Port       = $SMTPPort
        }
        if ($Attachments.Count -gt 0) {
            $resolved = @()
            foreach ($att in $Attachments) {
                if (Test-Path $att) {
                    $tmp = Join-Path $env:TEMP ("$(Split-Path $att -Leaf)_$(Get-Date -Format 'HHmmssff').tmp")
                    Copy-Item -Path $att -Destination $tmp -Force
                    $resolved   += $tmp
                    $tempCopies += $tmp
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
# Writes SQL to a temp file and executes via sqlplus -S "/ as sysdba".
# Returns hashtable: ExitCode, Output (string array).
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-SqlPlus {
    param(
        [string]$SqlBlock,
        [string]$SqlPlusExe
    )
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
    finally { Remove-Item $tmpSql, $tmpOut, $tmpErr -Force -ErrorAction SilentlyContinue }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Initialize-ExportDirectories   [PHASE 1]
# Creates physical filesystem directories and the Oracle DBA_DIRECTORY object.
# Returns $true on success, $false on failure.
# ─────────────────────────────────────────────────────────────────────────────
function Initialize-ExportDirectories {
    param(
        [PSCustomObject]$Config,
        [string]        $SqlPlusExe
    )

    Write-Log ('─' * 60) -Level INFO -Phase 'INIT'
    Write-Log 'PHASE 1 - Directory Initialisation' -Level INFO -Phase 'INIT'
    Write-Log ('─' * 60) -Level INFO -Phase 'INIT'

    # ── Step 1.1: Physical filesystem directories ─────────────────────────────
    Write-Log 'Step 1.1 : Filesystem directories' -Level INFO -Phase 'INIT'

    $dirsToCreate = [System.Collections.ArrayList]@()
    [void]$dirsToCreate.Add($Config.ExportBaseDirectory)
    [void]$dirsToCreate.Add($Config.LogDirectory)
    foreach ($schema in $Config.Schemas) {
        [void]$dirsToCreate.Add((Join-Path $Config.ExportBaseDirectory $schema.Name))
    }

    foreach ($dir in $dirsToCreate) {
        if (Test-Path $dir) {
            Write-Log "  EXISTS   : $dir" -Level INFO -Phase 'INIT'
        }
        else {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Write-Log "  CREATED  : $dir" -Level SUCCESS -Phase 'INIT'
        }
    }

    # ── Step 1.2: Oracle DBA_DIRECTORY object ─────────────────────────────────
    Write-Log 'Step 1.2 : Oracle DBA_DIRECTORY object' -Level INFO -Phase 'INIT'

    $dirObj  = $Config.OracleDirectoryObject
    $dirPath = $Config.ExportBaseDirectory
    $sqlPath = $dirPath.Replace("'", "''")

    $sqlBlock = @"
SET SERVEROUTPUT ON SIZE UNLIMITED
SET FEEDBACK OFF
SET HEADING OFF
WHENEVER SQLERROR EXIT SQL.SQLCODE

DECLARE
    v_exists NUMBER;
    v_action VARCHAR2(20);
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

    Write-Log "  Connecting to Oracle (/ as sysdba)..." -Level INFO -Phase 'INIT'
    $dbResult = Invoke-SqlPlus -SqlBlock $sqlBlock -SqlPlusExe $SqlPlusExe

    foreach ($line in $dbResult.Output) {
        if     ($line -match 'ERROR|ORA-') { Write-Log "  $line" -Level ERROR -Phase 'INIT' }
        elseif ($line.Trim() -ne '')        { Write-Log "  $line" -Level INFO  -Phase 'INIT' }
    }

    if ($dbResult.ExitCode -ne 0) {
        Write-Log "SQLPlus exited with code $($dbResult.ExitCode) - aborting." -Level ERROR -Phase 'INIT'
        return $false
    }

    Write-Log 'Phase 1 complete - all directories and Oracle DIRECTORY object are ready.' -Level SUCCESS -Phase 'INIT'
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: New-ExpdpParFile   [PHASE 2]
# Generates a schema-specific dated expdp parameter file.
# Returns hashtable: ParFilePath, DumpFile, LogFile, SchemaDir.
# ─────────────────────────────────────────────────────────────────────────────
function New-ExpdpParFile {
    param(
        [PSCustomObject]$SchemaConfig,
        [string]        $ExportBaseDir,
        [string]        $OracleDirectoryObject,
        [string]        $RunLabel,
        [string]        $RunDate
    )

    $schema      = $SchemaConfig.Name
    $schemaDir   = Join-Path $ExportBaseDir $schema
    if (-not (Test-Path $schemaDir)) { New-Item -ItemType Directory -Path $schemaDir | Out-Null }

    $dumpFile    = "${schema}_${RunDate}_${RunLabel}.dmp"
    $logFile     = "${schema}_${RunDate}_${RunLabel}_export.log"
    $parFilePath = Join-Path $schemaDir "${schema}_export.par"

    $lines = @(
        "USERID=""$($SchemaConfig.Credential)""",
        "SCHEMAS=$schema",
        "DIRECTORY=$OracleDirectoryObject",
        "DUMPFILE=$dumpFile",
        "LOGFILE=$logFile",
        "CLUSTER=N",
        "METRICS=YES",
        "EXCLUDE=STATISTICS"
    )

    if ($SchemaConfig.ExcludeTables -and $SchemaConfig.ExcludeTables.Count -gt 0) {
        $tableList = ($SchemaConfig.ExcludeTables | ForEach-Object { "'$_'" }) -join ','
        $lines    += "EXCLUDE=TABLE:`"IN ($tableList)`""
    }

    $lines | Set-Content -Path $parFilePath -Encoding ASCII
    Write-Log "Par file written: $parFilePath" -Level INFO -Phase 'EXPORT'

    return @{
        ParFilePath = $parFilePath
        DumpFile    = $dumpFile
        LogFile     = $logFile
        SchemaDir   = $schemaDir
    }
}



# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-SchemaExport   [PHASE 2]
# Executes expdp for one schema. Returns hashtable: Success, Duration, DumpSizeMB, Skipped.
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-SchemaExport {
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

    # ── Start notification ────────────────────────────────────────────────────
    $startBody = @"
<h3>Oracle Export Started</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Schema</b></td><td>$schema</td></tr>
  <tr><td><b>Started</b></td><td>$($startTime.ToString('yyyy-MM-dd HH:mm:ss'))</td></tr>
  <tr><td><b>Dump File</b></td><td>$($ParInfo.DumpFile)</td></tr>
  <tr><td><b>Mode</b></td><td>Oracle Standard Edition 2 - Sequential (no parallel)</td></tr>
  <tr><td><b>Dry Run</b></td><td>$IsDryRun</td></tr>
</table>
"@
    Send-ExportEmail -Subject "[EXPORT STARTED] $schema - Oracle Stage Refresh" `
                     -Body $startBody -To $allRecips `
                     -From $EmailConfig.From -SMTPServer $EmailConfig.SMTPServer -SMTPPort $EmailConfig.SMTPPort

    if ($IsDryRun) {
        Write-Log "[$schema] DRY RUN - expdp skipped." -Level WARN -Phase 'EXPORT'
        return @{ Success = $true; Duration = 0; DumpSizeMB = 0; Skipped = $true }
    }

    # ── Execute expdp with retry ──────────────────────────────────────────────
    $expdpExe  = Join-Path $OracleHome 'bin\expdp.exe'
    $stdoutTmp = Join-Path $ParInfo.SchemaDir 'expdp_stdout.tmp'
    $stderrTmp = Join-Path $ParInfo.SchemaDir 'expdp_stderr.tmp'
    $exitCode  = -1
    $attempt   = 0

    while ($attempt -le 2) {
        $attempt++
        Write-Log "[$schema] Attempt $attempt - executing expdp..." -Level INFO -Phase 'EXPORT'
        try {
            $proc = Start-Process -FilePath $expdpExe `
                                  -ArgumentList "parfile=`"$($ParInfo.ParFilePath)`"" `
                                  -WorkingDirectory $ParInfo.SchemaDir `
                                  -NoNewWindow -Wait -PassThru `
                                  -RedirectStandardOutput $stdoutTmp `
                                  -RedirectStandardError  $stderrTmp
            $exitCode = $proc.ExitCode
            break
        }
        catch {
            Write-Log "[$schema] expdp launch error (attempt $attempt): $_" -Level WARN -Phase 'EXPORT'
            if ($attempt -le 2) { Write-Log "[$schema] Retrying in 5 min..." -Level WARN -Phase 'EXPORT'; Start-Sleep -Seconds 300 }
        }
    }

    # ── Scan Oracle log for ORA- errors ──────────────────────────────────────
    $oracleLog   = Join-Path $ParInfo.SchemaDir $ParInfo.LogFile
    $oraErrors   = @()
    $hasOraError = $false
    if (Test-Path $oracleLog) {
        $oraErrors   = @(Select-String -Path $oracleLog -Pattern 'ORA-' | Select-Object -First 20)
        $hasOraError = $oraErrors.Count -gt 0
    }

    $endTime       = Get-Date
    $duration      = [math]::Round(($endTime - $startTime).TotalMinutes, 1)
    $exportBaseDir = Split-Path $ParInfo.SchemaDir -Parent
    $dumpPath      = Join-Path $exportBaseDir $ParInfo.DumpFile
    $dumpSizeMB    = 0
    if (Test-Path $dumpPath) { $dumpSizeMB = [math]::Round((Get-Item $dumpPath).Length / 1MB, 1) }
    Write-Log "[$schema] Dump: $dumpPath | Found: $(Test-Path $dumpPath) | Size: ${dumpSizeMB} MB" -Level INFO -Phase 'EXPORT'

    # ── Success ───────────────────────────────────────────────────────────────
    if ($exitCode -eq 0 -and -not $hasOraError) {
        Write-Log "[$schema] SUCCEEDED. Size: ${dumpSizeMB} MB | Duration: ${duration} min" -Level SUCCESS -Phase 'EXPORT'
        $successBody = @"
<h3 style='color:green'>Oracle Export Completed Successfully</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Schema</b></td><td>$schema</td></tr>
  <tr><td><b>Completed</b></td><td>$($endTime.ToString('yyyy-MM-dd HH:mm:ss'))</td></tr>
  <tr><td><b>Duration</b></td><td>${duration} minutes</td></tr>
  <tr><td><b>Dump File</b></td><td>$($ParInfo.DumpFile)</td></tr>
  <tr><td><b>Dump Path</b></td><td>$dumpPath</td></tr>
  <tr><td><b>Dump Size</b></td><td>${dumpSizeMB} MB</td></tr>
  <tr><td><b>Exit Code</b></td><td>$exitCode</td></tr>
</table>
"@
        Send-ExportEmail -Subject "[EXPORT COMPLETE] $schema - Oracle Stage Refresh" `
                         -Body $successBody -To $allRecips `
                         -From $EmailConfig.From -SMTPServer $EmailConfig.SMTPServer -SMTPPort $EmailConfig.SMTPPort
        return @{ Success = $true; Duration = $duration; DumpSizeMB = $dumpSizeMB; Skipped = $false }
    }

    # ── Failure ───────────────────────────────────────────────────────────────
    Write-Log "[$schema] FAILED. ExitCode=$exitCode | ORA-errors=$($oraErrors.Count)" -Level ERROR -Phase 'EXPORT'
    $oraDetail = if ($oraErrors.Count -gt 0) { ($oraErrors | ForEach-Object { $_.Line }) -join '<br/>' } else { 'None in Oracle log.' }
    $failBody  = @"
<h3 style='color:red'>ALERT: Oracle Export FAILED</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Schema</b></td><td>$schema</td></tr>
  <tr><td><b>Failed At</b></td><td>$($endTime.ToString('yyyy-MM-dd HH:mm:ss'))</td></tr>
  <tr><td><b>Duration</b></td><td>${duration} minutes</td></tr>
  <tr><td><b>Exit Code</b></td><td>$exitCode</td></tr>
  <tr><td><b>Attempts</b></td><td>$attempt</td></tr>
  <tr><td><b>ORA Errors</b></td><td>$oraDetail</td></tr>
</table>
<p>See attached logs for full details. The pipeline will continue with the next schema.</p>
"@
    $att = @()
    if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $att += $Script:LogFile }
    if (Test-Path $oracleLog)                              { $att += $oracleLog     }
    Send-ExportEmail -Subject "[EXPORT FAILED] $schema - Oracle Stage Refresh" `
                     -Body $failBody -To $dbaRecips `
                     -From $EmailConfig.From -SMTPServer $EmailConfig.SMTPServer -SMTPPort $EmailConfig.SMTPPort `
                     -Attachments $att
    return @{ Success = $false; Duration = $duration; DumpSizeMB = $dumpSizeMB; Skipped = $false }
}



# ═════════════════════════════════════════════════════════════════════════════
#  MAIN EXECUTION BLOCK
# ═════════════════════════════════════════════════════════════════════════════
try {
    # ── Load config ───────────────────────────────────────────────────────────
    if (-not (Test-Path $ConfigPath)) { throw "Config file not found: $ConfigPath" }
    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json

    # ── Oracle environment ────────────────────────────────────────────────────
    $env:ORACLE_HOME = $config.OracleHome
    $env:PATH        = "$($config.OracleHome)\bin;$($env:PATH)"
    $sqlplusExe      = Join-Path $config.OracleHome 'bin\sqlplus.exe'
    if (-not (Test-Path $sqlplusExe)) { throw "sqlplus.exe not found at: $sqlplusExe" }

    # ── Open log file (ensure log dir exists first) ───────────────────────────
    if (-not (Test-Path $config.LogDirectory)) {
        New-Item -ItemType Directory -Path $config.LogDirectory -Force | Out-Null
    }
    $Script:LogFile   = Join-Path $config.LogDirectory "OracleExport_${RunTimestamp}.log"
    $Script:LogWriter = [System.IO.StreamWriter]::new($Script:LogFile, $true, [System.Text.Encoding]::UTF8)
    $Script:LogWriter.AutoFlush = $true

    # ── Startup banner ────────────────────────────────────────────────────────
    Write-Log ('=' * 65) -Level INFO -Phase 'STARTUP'
    Write-Log 'Oracle Stage Refresh - Full Export Pipeline'        -Level INFO -Phase 'STARTUP'
    Write-Log "Server    : $($env:COMPUTERNAME)"                   -Level INFO -Phase 'STARTUP'
    Write-Log "Run ID    : $RunTimestamp"                          -Level INFO -Phase 'STARTUP'
    Write-Log "Log File  : $($Script:LogFile)"                     -Level INFO -Phase 'STARTUP'
    Write-Log "Oracle    : $($config.OracleHome)"                  -Level INFO -Phase 'STARTUP'
    Write-Log "Export Dir: $($config.ExportBaseDirectory)"         -Level INFO -Phase 'STARTUP'
    Write-Log "Schemas   : $(($config.Schemas | ForEach-Object { $_.Name }) -join ', ')" -Level INFO -Phase 'STARTUP'
    Write-Log "SkipInit  : $($SkipInit.IsPresent)"                 -Level INFO -Phase 'STARTUP'
    Write-Log "InitOnly  : $($InitOnly.IsPresent)"                 -Level INFO -Phase 'STARTUP'
    Write-Log "DryRun    : $($DryRun.IsPresent)"                   -Level INFO -Phase 'STARTUP'
    Write-Log ('=' * 65) -Level INFO -Phase 'STARTUP'

    # ══════════════════════════════════════════════════════════════════════════
    #  PHASE 1 — Directory Initialisation
    # ══════════════════════════════════════════════════════════════════════════
    if ($SkipInit.IsPresent) {
        Write-Log 'Phase 1 skipped (-SkipInit specified).' -Level WARN -Phase 'INIT'
    }
    else {
        $initOk = Initialize-ExportDirectories -Config $config -SqlPlusExe $sqlplusExe
        if (-not $initOk) {
            throw 'Phase 1 (Directory Initialisation) failed. Export pipeline will not run.'
        }
    }

    # Exit here if -InitOnly was requested
    if ($InitOnly.IsPresent) {
        Write-Log 'InitOnly mode — exiting after Phase 1.' -Level INFO -Phase 'INIT'
        exit 0
    }

    # ══════════════════════════════════════════════════════════════════════════
    #  PHASE 2 — Export Pipeline
    # ══════════════════════════════════════════════════════════════════════════
    Write-Log '' -Level INFO -Phase 'EXPORT'
    Write-Log ('─' * 60) -Level INFO -Phase 'EXPORT'
    Write-Log 'PHASE 2 - Export Pipeline' -Level INFO -Phase 'EXPORT'
    Write-Log ('─' * 60) -Level INFO -Phase 'EXPORT'

    $RunDate   = Get-Date -Format 'MMMdd_yyyy'
    $allRecips = @($config.Email.DBARecipients) + @($config.Email.BusinessRecipients)
    $results   = [System.Collections.ArrayList]@()

    # Pipeline start email
    $schemaList   = ($config.Schemas | ForEach-Object { $_.Name }) -join ', '
    $pipelineBody = @"
<h2>Oracle Stage Database Refresh &mdash; Export Pipeline Started</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Schemas</b></td><td>$schemaList</td></tr>
  <tr><td><b>Oracle Directory</b></td><td>$($config.OracleDirectoryObject)</td></tr>
  <tr><td><b>Export Base Dir</b></td><td>$($config.ExportBaseDirectory)</td></tr>
  <tr><td><b>Dry Run</b></td><td>$($DryRun.IsPresent)</td></tr>
  <tr><td><b>Started</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
<p><i>You will receive individual emails for each schema start and completion.</i></p>
"@
    Send-ExportEmail -Subject "[EXPORT PIPELINE STARTED] Oracle Stage Refresh - $RunTimestamp" `
                     -Body $pipelineBody -To $allRecips `
                     -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort

    # Process each schema sequentially
    foreach ($schemaConfig in $config.Schemas) {
        Write-Log "--- Schema: $($schemaConfig.Name) ---" -Level INFO -Phase 'EXPORT'

        $parInfo = New-ExpdpParFile -SchemaConfig          $schemaConfig `
                                    -ExportBaseDir         $config.ExportBaseDirectory `
                                    -OracleDirectoryObject $config.OracleDirectoryObject `
                                    -RunLabel              $config.RunLabel `
                                    -RunDate               $RunDate

        $result = Invoke-SchemaExport -SchemaConfig $schemaConfig `
                                      -ParInfo      $parInfo `
                                      -OracleHome   $config.OracleHome `
                                      -EmailConfig  $config.Email `
                                      -IsDryRun     $DryRun.IsPresent

        [void]$results.Add([PSCustomObject]@{
            Schema     = $schemaConfig.Name
            Success    = $result.Success
            DumpSizeMB = $result.DumpSizeMB
            Duration   = $result.Duration
            Skipped    = $result.Skipped
            DumpFile   = $parInfo.DumpFile
        })

        if (-not $result.Success) {
            Write-Log "[$($schemaConfig.Name)] Failed - continuing to next schema." -Level WARN -Phase 'EXPORT'
        }
    }

    # ── Final summary ─────────────────────────────────────────────────────────
    $successCount = @($results | Where-Object {  $_.Success }).Count
    $failCount    = @($results | Where-Object { -not $_.Success }).Count
    $totalSizeMB  = [math]::Round(($results | Measure-Object -Property DumpSizeMB -Sum).Sum, 1)
    $totalMins    = [math]::Round(($results | Measure-Object -Property Duration   -Sum).Sum, 1)

    $tableRows = $results | ForEach-Object {
        $color  = if ($_.Success) { 'green' } else { 'red' }
        $status = if ($_.Skipped) { 'SKIPPED (DryRun)' } elseif ($_.Success) { 'SUCCESS' } else { 'FAILED' }
        "<tr><td>$($_.Schema)</td><td style='color:$color'><b>$status</b></td><td>$($_.DumpSizeMB) MB</td><td>$($_.Duration) min</td><td>$($_.DumpFile)</td></tr>"
    }

    $overallStatus = if ($failCount -gt 0) { 'COMPLETED WITH ERRORS' } else { 'COMPLETED SUCCESSFULLY' }
    $summaryBody   = @"
<h2>Oracle Stage Database Refresh &mdash; Export Pipeline Summary</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Total Schemas</b></td><td>$($config.Schemas.Count)</td></tr>
  <tr><td><b>Successful</b></td><td style='color:green'><b>$successCount</b></td></tr>
  <tr><td><b>Failed</b></td><td style='color:red'><b>$failCount</b></td></tr>
  <tr><td><b>Total Dump Size</b></td><td>${totalSizeMB} MB</td></tr>
  <tr><td><b>Total Duration</b></td><td>${totalMins} minutes</td></tr>
</table><br/>
<h3>Per-Schema Results</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr style='background:#ddd'><th>Schema</th><th>Status</th><th>Size</th><th>Duration</th><th>Dump File</th></tr>
  $($tableRows -join '')
</table><br/><p>Full log attached.</p>
"@
    Write-Log "Pipeline finished. Success=$successCount | Failed=$failCount | Total=${totalSizeMB} MB | ${totalMins} min" -Level SUCCESS -Phase 'SUMMARY'

    $summaryAtt = @(); if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $summaryAtt += $Script:LogFile }
    Send-ExportEmail -Subject "[EXPORT SUMMARY] $overallStatus - Oracle Stage Refresh $RunTimestamp" `
                     -Body $summaryBody -To $allRecips `
                     -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort `
                     -Attachments $summaryAtt

    if ($failCount -gt 0) { exit 1 }

    # ══ PHASE 3 — Tablespace Snapshot ════════════════════════════════════════
    Write-Log ('─' * 60) -Level INFO -Phase 'SNAPSHOT'
    Write-Log 'PHASE 3 - Capturing Production tablespace snapshot...' -Level INFO -Phase 'SNAPSHOT'
    $snapshotScript = Join-Path $PSScriptRoot 'New-TablespaceSnapshot.ps1'
    if (Test-Path $snapshotScript) {
        try {
            & $snapshotScript -ConfigPath $ConfigPath
            Write-Log 'Tablespace snapshot complete.' -Level SUCCESS -Phase 'SNAPSHOT'
        }
        catch {
            Write-Log "Tablespace snapshot failed (non-fatal): $_" -Level WARN -Phase 'SNAPSHOT'
        }
    }
    else {
        Write-Log "Snapshot script not found at $snapshotScript — skipping." -Level WARN -Phase 'SNAPSHOT'
    }
}
catch {
    $errMsg = $_.Exception.Message
    Write-Log "CRITICAL: $errMsg" -Level ERROR -Phase 'FATAL'
    $critBody = @"
<h2 style='color:red'>CRITICAL: Oracle Export Pipeline Stopped</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Error</b></td><td>$errMsg</td></tr>
  <tr><td><b>Time</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
<p>Please review the attached log immediately.</p>
"@
    try {
        if ($null -ne $config -and $null -ne $config.Email) {
            $ca = @(); if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $ca += $Script:LogFile }
            Send-ExportEmail -Subject "[CRITICAL ERROR] Oracle Export - $RunTimestamp" `
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
