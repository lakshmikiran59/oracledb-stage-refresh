<#
.SYNOPSIS
    Oracle Production Schema Export Automation Script (Story 1 - DBA-555).

.DESCRIPTION
    Automates Oracle Data Pump Export (expdp) for all 8 agreed schemas:
    ASHLEY, CIM, DIS, EMANIFEST, FINANCE, LEAN, SQLORACLE, TAPLSQL.

    For each schema the script will:
      - Generate a schema-specific dated .par parameter file
      - Execute expdp sequentially (Oracle Standard Edition 2 - no parallel)
      - Send email notifications on start and completion
      - Send alert emails with log attachments on failure
      - Produce a final HTML summary report

.PARAMETER ConfigPath
    Path to ExportConfig.json. Defaults to .\Config\ExportConfig.json.

.PARAMETER DryRun
    Validate configuration and generate .par files only. Does NOT execute expdp.

.EXAMPLE
    .\Invoke-OracleExport.ps1
    .\Invoke-OracleExport.ps1 -DryRun
    .\Invoke-OracleExport.ps1 -ConfigPath "D:\Custom\ExportConfig.json"

.NOTES
    Compatible  : PowerShell 5.0+
    Oracle Ed.  : Standard Edition 2 (CLUSTER=N, no PARALLEL, no COMPRESSION)
    Jira        : DBA-555 | Story 1 - Automated Export Module
    Author      : DBA Team - Ashley Furniture India
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "$PSScriptRoot\Config\ExportConfig.json",
    [switch]$DryRun
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Script-scoped variables set during initialisation
$Script:LogFile    = $null
$Script:LogWriter  = $null   # Persistent StreamWriter - avoids repeated open/close file locking
$RunTimestamp      = Get-Date -Format 'yyyyMMdd_HHmmss'

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Write-ExportLog
# Writes a timestamped, colour-coded entry to console and the run log file.
# ─────────────────────────────────────────────────────────────────────────────
function Write-ExportLog {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO',
        [string]$Phase = 'EXPORT'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] [$Phase] $Message"

    switch ($Level) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan    }
        'WARN'    { Write-Host $line -ForegroundColor Yellow  }
        'ERROR'   { Write-Host $line -ForegroundColor Red     }
        'SUCCESS' { Write-Host $line -ForegroundColor Green   }
    }
    # Use the persistent StreamWriter so the file handle is never closed mid-run.
    # This prevents "file in use" errors when Send-MailMessage opens the same file
    # as an attachment and fails to release the handle on SMTP error.
    if ($Script:LogWriter) {
        $Script:LogWriter.WriteLine($line)
        $Script:LogWriter.Flush()
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Send-ExportEmail
# Sends an HTML email via Send-MailMessage (PowerShell 5.0 built-in).
# Silently logs a warning on failure - does not abort the pipeline.
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

        # Copy each attachment to a temp file before passing to Send-MailMessage.
        # Send-MailMessage does NOT release file handles on SMTP failure, which locks
        # the live log file and causes subsequent Add-Content / StreamWriter calls to fail.
        $tempCopies = @()
        if ($Attachments.Count -gt 0) {
            $resolvedAttachments = @()
            foreach ($att in $Attachments) {
                if (Test-Path $att) {
                    $tmpCopy = Join-Path $env:TEMP ("$(Split-Path $att -Leaf)_$(Get-Date -Format 'HHmmssff').tmp")
                    Copy-Item -Path $att -Destination $tmpCopy -Force
                    $resolvedAttachments += $tmpCopy
                    $tempCopies         += $tmpCopy
                }
            }
            if ($resolvedAttachments.Count -gt 0) { $params['Attachments'] = $resolvedAttachments }
        }

        Send-MailMessage @params
        Write-ExportLog "Email sent: $Subject" -Level INFO
    }
    catch {
        Write-ExportLog "Email send failed (non-fatal): $_" -Level WARN
    }
    finally {
        # Always clean up temp attachment copies regardless of success or failure
        if ($tempCopies.Count -gt 0) {
            $tempCopies | Remove-Item -Force -ErrorAction SilentlyContinue
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: New-ExpdpParFile
# Generates a schema-specific expdp parameter file with date-stamped filenames.
# Returns a hashtable with par file path, dump file name, log file name.
# ─────────────────────────────────────────────────────────────────────────────
function New-ExpdpParFile {
    param(
        [PSCustomObject] $SchemaConfig,
        [string]         $ExportBaseDir,
        [string]         $OracleDirectoryObject,
        [string]         $RunLabel,
        [string]         $RunDate
    )

    $schema    = $SchemaConfig.Name
    $schemaDir = Join-Path $ExportBaseDir $schema
    if (-not (Test-Path $schemaDir)) {
        New-Item -ItemType Directory -Path $schemaDir | Out-Null
    }

    $dumpFile    = "${schema}_${RunDate}_${RunLabel}.dmp"
    $logFile     = "${schema}_${RunDate}_${RunLabel}_export.log"
    $parFilePath = Join-Path $schemaDir "${schema}_export.par"

    # Core parameter lines - Oracle SE2 compatible (no PARALLEL, no COMPRESSION)
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

    # Append schema-specific table exclusions when defined
    if ($SchemaConfig.ExcludeTables -and $SchemaConfig.ExcludeTables.Count -gt 0) {
        $tableList = ($SchemaConfig.ExcludeTables | ForEach-Object { "'$_'" }) -join ','
        $lines    += "EXCLUDE=TABLE:`"IN ($tableList)`""
    }

    $lines | Set-Content -Path $parFilePath -Encoding ASCII
    Write-ExportLog "Par file written: $parFilePath" -Level INFO -Phase 'EXPORT'

    return @{
        ParFilePath = $parFilePath
        DumpFile    = $dumpFile
        LogFile     = $logFile
        SchemaDir   = $schemaDir
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-SchemaExport
# Executes expdp for a single schema, with start/completion/failure emails.
# Returns a hashtable: Success, Duration (min), DumpSizeMB, Skipped.
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-SchemaExport {
    param(
        [PSCustomObject] $SchemaConfig,
        [hashtable]      $ParInfo,
        [string]         $OracleHome,
        [PSCustomObject] $EmailConfig,
        [bool]           $IsDryRun
    )

    $schema     = $SchemaConfig.Name
    $startTime  = Get-Date
    $allRecips  = @($EmailConfig.DBARecipients) + @($EmailConfig.BusinessRecipients)
    $dbaRecips  = @($EmailConfig.DBARecipients)

    # ── Export START notification ────────────────────────────────────────────
    $startBody = @"
<h3>Oracle Export Started</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Schema</b></td><td>$schema</td></tr>
  <tr><td><b>Started</b></td><td>$($startTime.ToString('yyyy-MM-dd HH:mm:ss'))</td></tr>
  <tr><td><b>Dump File</b></td><td>$($ParInfo.DumpFile)</td></tr>
  <tr><td><b>Oracle Directory</b></td><td>$($SchemaConfig.TNSAlias)</td></tr>
  <tr><td><b>Mode</b></td><td>Oracle Standard Edition 2 - Sequential (no parallel)</td></tr>
  <tr><td><b>Dry Run</b></td><td>$IsDryRun</td></tr>
</table>
"@
    Send-ExportEmail -Subject "[EXPORT STARTED] $schema - Oracle Stage Refresh" `
                     -Body $startBody -To $allRecips `
                     -From $EmailConfig.From -SMTPServer $EmailConfig.SMTPServer -SMTPPort $EmailConfig.SMTPPort

    # ── Dry Run guard ────────────────────────────────────────────────────────
    if ($IsDryRun) {
        Write-ExportLog "[$schema] DRY RUN - expdp execution skipped." -Level WARN -Phase 'EXPORT'
        return @{ Success = $true; Duration = 0; DumpSizeMB = 0; Skipped = $true }
    }

    # ── Execute expdp with retry on transient errors ─────────────────────────
    $expdpExe   = Join-Path $OracleHome 'bin\expdp.exe'
    $parFile    = $ParInfo.ParFilePath
    $oracleLog  = Join-Path $ParInfo.SchemaDir $ParInfo.LogFile
    $stdoutTmp  = Join-Path $ParInfo.SchemaDir 'expdp_stdout.tmp'
    $stderrTmp  = Join-Path $ParInfo.SchemaDir 'expdp_stderr.tmp'
    $exitCode   = -1
    $attempt    = 0
    $maxRetries = 2

    while ($attempt -le $maxRetries) {
        $attempt++
        Write-ExportLog "[$schema] Attempt $attempt - executing expdp..." -Level INFO -Phase 'EXPORT'
        try {
            $proc = Start-Process -FilePath $expdpExe `
                                  -ArgumentList "parfile=`"$parFile`"" `
                                  -WorkingDirectory $ParInfo.SchemaDir `
                                  -NoNewWindow -Wait -PassThru `
                                  -RedirectStandardOutput $stdoutTmp `
                                  -RedirectStandardError  $stderrTmp
            $exitCode = $proc.ExitCode
            break   # clean exit from loop
        }
        catch {
            Write-ExportLog "[$schema] expdp launch error (attempt $attempt): $_" -Level WARN -Phase 'EXPORT'
            if ($attempt -le $maxRetries) {
                Write-ExportLog "[$schema] Retrying in 5 minutes..." -Level WARN -Phase 'EXPORT'
                Start-Sleep -Seconds 300
            }
        }
    }

    # ── Scan Oracle log for ORA- errors (exit code 0 can still have warnings) ─
    $oraErrors   = @()
    $hasOraError = $false
    if (Test-Path $oracleLog) {
        $oraErrors   = @(Select-String -Path $oracleLog -Pattern 'ORA-' | Select-Object -First 20)
        $hasOraError = $oraErrors.Count -gt 0
    }

    $endTime    = Get-Date
    $duration   = [math]::Round(($endTime - $startTime).TotalMinutes, 1)

    # Oracle writes the dump file to the DIRECTORY object root (e.g. E:\Datapump),
    # NOT to the schema subfolder (E:\Datapump\ASHLEY). Derive base dir from SchemaDir parent.
    $exportBaseDir = Split-Path $ParInfo.SchemaDir -Parent
    $dumpPath      = Join-Path $exportBaseDir $ParInfo.DumpFile
    $dumpSizeMB    = 0
    if (Test-Path $dumpPath) {
        $dumpSizeMB = [math]::Round((Get-Item $dumpPath).Length / 1MB, 1)
    }
    Write-ExportLog "[$schema] Dump file path: $dumpPath | Found: $(Test-Path $dumpPath) | Size: ${dumpSizeMB} MB" -Level INFO -Phase 'EXPORT'

    # ── SUCCESS path ─────────────────────────────────────────────────────────
    if ($exitCode -eq 0 -and -not $hasOraError) {
        Write-ExportLog "[$schema] Export SUCCEEDED. Size: ${dumpSizeMB} MB | Duration: ${duration} min" -Level SUCCESS -Phase 'EXPORT'

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

    # ── FAILURE path ─────────────────────────────────────────────────────────
    Write-ExportLog "[$schema] Export FAILED. ExitCode=$exitCode | ORA-errors=$($oraErrors.Count)" -Level ERROR -Phase 'EXPORT'

    $oraDetail  = if ($oraErrors.Count -gt 0) {
                      ($oraErrors | ForEach-Object { $_.Line }) -join '<br/>'
                  } else { 'None detected in Oracle log.' }

    $failBody = @"
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
    $attachments = @()
    if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $attachments += $Script:LogFile }
    if (Test-Path $oracleLog)                              { $attachments += $oracleLog     }

    Send-ExportEmail -Subject "[EXPORT FAILED] $schema - Oracle Stage Refresh" `
                     -Body $failBody -To $dbaRecips `
                     -From $EmailConfig.From -SMTPServer $EmailConfig.SMTPServer -SMTPPort $EmailConfig.SMTPPort `
                     -Attachments $attachments

    return @{ Success = $false; Duration = $duration; DumpSizeMB = $dumpSizeMB; Skipped = $false }
}


# ═════════════════════════════════════════════════════════════════════════════
#  MAIN EXECUTION BLOCK
# ═════════════════════════════════════════════════════════════════════════════
try {

    # ── Load and validate configuration ──────────────────────────────────────
    if (-not (Test-Path $ConfigPath)) {
        throw "Configuration file not found: $ConfigPath"
    }
    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json

    # ── Initialise log directory and run log file ─────────────────────────────
    if (-not (Test-Path $config.LogDirectory)) {
        New-Item -ItemType Directory -Path $config.LogDirectory | Out-Null
    }
    $Script:LogFile = Join-Path $config.LogDirectory "OracleExport_${RunTimestamp}.log"

    # Open a persistent StreamWriter with FileShare.Read so:
    #   - We hold the only write handle for the life of the script
    #   - Other processes (email clients, log viewers) can still read the file
    #   - No repeated open/close races from Add-Content
    $Script:LogWriter = [System.IO.StreamWriter]::new(
        $Script:LogFile,
        $true,   # append mode
        [System.Text.Encoding]::UTF8
    )
    $Script:LogWriter.AutoFlush = $true

    # ── Set Oracle environment ────────────────────────────────────────────────
    $env:ORACLE_HOME = $config.OracleHome
    $env:PATH        = "$($config.OracleHome)\bin;$($env:PATH)"

    $RunDate   = Get-Date -Format 'MMMdd_yyyy'
    $allRecips = @($config.Email.DBARecipients) + @($config.Email.BusinessRecipients)
    $results   = [System.Collections.ArrayList]@()

    Write-ExportLog ('=' * 65) -Level INFO -Phase 'STARTUP'
    Write-ExportLog 'Oracle Stage Refresh - Export Pipeline Starting'    -Level INFO -Phase 'STARTUP'
    Write-ExportLog "Server    : $($env:COMPUTERNAME)"                   -Level INFO -Phase 'STARTUP'
    Write-ExportLog "Run ID    : $RunTimestamp"                          -Level INFO -Phase 'STARTUP'
    Write-ExportLog "Log File  : $($Script:LogFile)"                     -Level INFO -Phase 'STARTUP'
    Write-ExportLog "Oracle    : $($config.OracleHome)"                  -Level INFO -Phase 'STARTUP'
    Write-ExportLog "Export Dir: $($config.ExportBaseDirectory)"         -Level INFO -Phase 'STARTUP'
    Write-ExportLog "Dry Run   : $($DryRun.IsPresent)"                   -Level INFO -Phase 'STARTUP'
    Write-ExportLog "Schemas   : $(($config.Schemas | ForEach-Object { $_.Name }) -join ', ')" -Level INFO -Phase 'STARTUP'
    Write-ExportLog ('=' * 65) -Level INFO -Phase 'STARTUP'

    # ── Pipeline start email to all stakeholders ──────────────────────────────
    $schemaList   = ($config.Schemas | ForEach-Object { $_.Name }) -join ', '
    $pipelineBody = @"
<h2>Oracle Stage Database Refresh &mdash; Export Pipeline Started</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Schemas</b></td><td>$schemaList</td></tr>
  <tr><td><b>Oracle Directory</b></td><td>$($config.OracleDirectoryObject)</td></tr>
  <tr><td><b>Export Base Dir</b></td><td>$($config.ExportBaseDirectory)</td></tr>
  <tr><td><b>Dry Run Mode</b></td><td>$($DryRun.IsPresent)</td></tr>
  <tr><td><b>Started</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
<p><i>You will receive individual emails for each schema export start and completion.</i></p>
"@
    Send-ExportEmail -Subject "[EXPORT PIPELINE STARTED] Oracle Stage Refresh - $RunTimestamp" `
                     -Body $pipelineBody -To $allRecips `
                     -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort

    # ── Process each schema sequentially (Oracle SE2 - no parallel) ───────────
    foreach ($schemaConfig in $config.Schemas) {
        $schema = $schemaConfig.Name
        Write-ExportLog "--- Processing schema: $schema ---" -Level INFO -Phase 'EXPORT'

        $parInfo = New-ExpdpParFile -SchemaConfig          $schemaConfig `
                                    -ExportBaseDir         $config.ExportBaseDirectory `
                                    -OracleDirectoryObject $config.OracleDirectoryObject `
                                    -RunLabel              $config.RunLabel `
                                    -RunDate               $RunDate

        $result  = Invoke-SchemaExport -SchemaConfig $schemaConfig `
                                       -ParInfo      $parInfo `
                                       -OracleHome   $config.OracleHome `
                                       -EmailConfig  $config.Email `
                                       -IsDryRun     $DryRun.IsPresent

        [void]$results.Add([PSCustomObject]@{
            Schema     = $schema
            Success    = $result.Success
            DumpSizeMB = $result.DumpSizeMB
            Duration   = $result.Duration
            Skipped    = $result.Skipped
            DumpFile   = $parInfo.DumpFile
        })

        if (-not $result.Success) {
            Write-ExportLog "[$schema] Failed - continuing to next schema." -Level WARN -Phase 'EXPORT'
        }
    }

    # ── Build and send final summary report ───────────────────────────────────
    $successCount = @($results | Where-Object { $_.Success  }).Count
    $failCount    = @($results | Where-Object { -not $_.Success }).Count
    $totalSizeMB  = [math]::Round(($results | Measure-Object -Property DumpSizeMB -Sum).Sum, 1)
    $totalMins    = [math]::Round(($results | Measure-Object -Property Duration   -Sum).Sum, 1)

    $tableRows = $results | ForEach-Object {
        $color  = if ($_.Success) { 'green' } else { 'red' }
        $status = if ($_.Skipped) { 'SKIPPED (DryRun)' } elseif ($_.Success) { 'SUCCESS' } else { 'FAILED' }
        "<tr>
           <td>$($_.Schema)</td>
           <td style='color:$color'><b>$status</b></td>
           <td>$($_.DumpSizeMB) MB</td>
           <td>$($_.Duration) min</td>
           <td>$($_.DumpFile)</td>
         </tr>"
    }

    $summaryBody = @"
<h2>Oracle Stage Database Refresh &mdash; Export Pipeline Summary</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Total Schemas</b></td><td>$($config.Schemas.Count)</td></tr>
  <tr><td><b>Successful</b></td><td style='color:green'><b>$successCount</b></td></tr>
  <tr><td><b>Failed</b></td><td style='color:red'><b>$failCount</b></td></tr>
  <tr><td><b>Total Dump Size</b></td><td>${totalSizeMB} MB</td></tr>
  <tr><td><b>Total Duration</b></td><td>${totalMins} minutes</td></tr>
</table>
<br/>
<h3>Per-Schema Results</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr style='background:#ddd'><th>Schema</th><th>Status</th><th>Size</th><th>Duration</th><th>Dump File</th></tr>
  $($tableRows -join "`n")
</table>
<br/><p>Full log attached for review.</p>
"@

    $overallStatus = if ($failCount -gt 0) { 'COMPLETED WITH ERRORS' } else { 'COMPLETED SUCCESSFULLY' }
    Write-ExportLog "Export pipeline finished. Success=$successCount | Failed=$failCount | Total=${totalSizeMB} MB | ${totalMins} min" `
                    -Level SUCCESS -Phase 'SUMMARY'

    $summaryAttachments = @()
    if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $summaryAttachments += $Script:LogFile }

    Send-ExportEmail -Subject "[EXPORT SUMMARY] $overallStatus - Oracle Stage Refresh $RunTimestamp" `
                     -Body $summaryBody -To $allRecips `
                     -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort `
                     -Attachments $summaryAttachments

    # Return non-zero exit code if any schema failed, so Task Scheduler can detect failure
    if ($failCount -gt 0) { exit 1 }

}
catch {
    # ── Unhandled critical error ───────────────────────────────────────────────
    $errMsg = $_.Exception.Message
    Write-ExportLog "CRITICAL UNHANDLED ERROR: $errMsg" -Level ERROR -Phase 'FATAL'


    $critBody = @"
<h2 style='color:red'>CRITICAL: Oracle Export Pipeline Stopped Unexpectedly</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Error</b></td><td>$errMsg</td></tr>
  <tr><td><b>Time</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
<p>Please review the attached log immediately. No further schemas were processed.</p>
"@
    try {
        if ($null -ne $config -and $null -ne $config.Email) {
            $critAttach = @()
            if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $critAttach += $Script:LogFile }
            Send-ExportEmail -Subject "[CRITICAL ERROR] Oracle Export Pipeline - $RunTimestamp" `
                             -Body $critBody -To @($config.Email.DBARecipients) `
                             -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort `
                             -Attachments $critAttach
        }
    }
    catch { Write-Host "Failed to send critical error email: $_" -ForegroundColor Red }

    exit 1
}
finally {
    # ── Always runs: close the StreamWriter so the log file is properly flushed
    # and released, whether the pipeline succeeded, failed, or was interrupted.
    if ($null -ne $Script:LogWriter) {
        try   { $Script:LogWriter.Close(); $Script:LogWriter.Dispose() }
        catch { Write-Host "Warning: could not close log writer: $_" -ForegroundColor Yellow }
        $Script:LogWriter = $null
    }
}
