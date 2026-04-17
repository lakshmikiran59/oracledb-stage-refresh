<#
.SYNOPSIS
    Oracle Stage Refresh — DMP File Transfer (Story 2 / DBA-556).

.DESCRIPTION
    Copies all schema .dmp export files from the Production export directory to the
    Stage import directory over UNC using Robocopy (restartable, retry-capable).

    For each schema the script will:
      - Locate the most-recent .dmp file in the source directory for that schema
      - Robocopy it to \\<StageServer>\f$\DataImportFromProd\  (created if absent)
      - Verify destination file size matches source
      - Send HTML email notifications on start, per-file completion, and final summary

    Source files are NOT deleted after transfer.

.PARAMETER StageServer
    Hostname of the target Stage server.  e.g.  aazeud-oracle-stg

.PARAMETER ConfigPath
    Path to ExportConfig.json. Defaults to ..\Export\Config\ExportConfig.json.

.PARAMETER DryRun
    Shows what would be transferred but does NOT copy any files.

.PARAMETER TargetBasePath
    Override the target UNC root.  Default: \\<StageServer>\f$\DataImportFromProd

.EXAMPLE
    # Standard transfer to stage server
    .\Start-DataTransfer.ps1 -StageServer aazeud-oracle-stg

    # Preview without copying
    .\Start-DataTransfer.ps1 -StageServer aazeud-oracle-stg -DryRun

    # Custom target path
    .\Start-DataTransfer.ps1 -StageServer aazeud-oracle-stg -TargetBasePath "\\aazeud-oracle-stg\e$\Import"

.NOTES
    Compatible : PowerShell 5.0+
    Requires   : Robocopy (built-in on Windows Server 2008+)
                 Network access to StageServer admin share (f$)
    Jira       : DBA-556 | Story 2 - Data Transfer & File Preparation
    Author     : DBA Team - Ashley Furniture India
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$StageServer,

    [string]$ConfigPath = "$PSScriptRoot\..\Export\Config\ExportConfig.json",

    [switch]$DryRun,

    [string]$TargetBasePath = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$Script:LogFile   = $null
$Script:LogWriter = $null
$RunTimestamp     = Get-Date -Format 'yyyyMMdd_HHmmss'

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Write-Log
# Timestamped colour-coded logger — console + persistent StreamWriter file log.
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO',
        [string]$Phase = 'TRANSFER'
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
# FUNCTION: Send-TransferEmail
# HTML email via Send-MailMessage (PS 5.0). Copies attachments to temp files
# to prevent handle-lock errors on SMTP failure.
# ─────────────────────────────────────────────────────────────────────────────
function Send-TransferEmail {
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
# FUNCTION: Invoke-FileTransfer
# Robocopy a single .dmp file from source dir to target UNC dir.
# Returns hashtable: Success, FileSizeMB, Duration, SourcePath, DestPath.
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-FileTransfer {
    param(
        [string]$Schema,
        [string]$SourceFile,      # Full path to source .dmp file
        [string]$TargetDir,       # UNC destination directory
        [bool]  $IsDryRun
    )

    $startTime    = Get-Date
    $fileName     = Split-Path $SourceFile -Leaf
    $destPath     = Join-Path $TargetDir $fileName
    $sourceSizeMB = [math]::Round((Get-Item $SourceFile).Length / 1MB, 1)

    Write-Log "[$Schema] Source : $SourceFile ($sourceSizeMB MB)" -Level INFO
    Write-Log "[$Schema] Target : $destPath"                      -Level INFO

    if ($IsDryRun) {
        Write-Log "[$Schema] DRY RUN - transfer skipped." -Level WARN
        return @{ Success = $true; FileSizeMB = $sourceSizeMB; Duration = 0
                  SourcePath = $SourceFile; DestPath = $destPath; Skipped = $true }
    }

    # ── Copy-Item transfer ────────────────────────────────────────────────────
    Write-Log "[$Schema] Copying $sourceSizeMB MB -> $destPath ..." -Level INFO
    try {
        Copy-Item -Path $SourceFile -Destination $destPath -Force
    }
    catch {
        $duration = [math]::Round(((Get-Date) - $startTime).TotalMinutes, 1)
        Write-Log "[$Schema] Copy-Item FAILED: $_" -Level ERROR
        return @{ Success = $false; FileSizeMB = $sourceSizeMB; Duration = $duration
                  SourcePath = $SourceFile; DestPath = $destPath; Skipped = $false }
    }

    # Calculate duration here so it is always defined on the success path
    $duration = [math]::Round(((Get-Date) - $startTime).TotalMinutes, 1)

    # ── Verify destination file size matches source ───────────────────────────
    if (-not (Test-Path $destPath)) {
        Write-Log "[$Schema] Destination file not found after copy: $destPath" -Level ERROR
        return @{ Success = $false; FileSizeMB = $sourceSizeMB; Duration = $duration
                  SourcePath = $SourceFile; DestPath = $destPath; Skipped = $false }
    }

    $destSizeMB = [math]::Round((Get-Item $destPath).Length / 1MB, 1)
    Write-Log "[$Schema] Source: ${sourceSizeMB} MB  |  Dest: ${destSizeMB} MB" -Level INFO

    if ($destSizeMB -ne $sourceSizeMB) {
        Write-Log "[$Schema] SIZE MISMATCH - transfer may be incomplete!" -Level ERROR
        return @{ Success = $false; FileSizeMB = $destSizeMB; Duration = $duration
                  SourcePath = $SourceFile; DestPath = $destPath; Skipped = $false }
    }

    Write-Log "[$Schema] Transfer SUCCEEDED. ${destSizeMB} MB in ${duration} min." -Level SUCCESS
    return @{ Success = $true; FileSizeMB = $destSizeMB; Duration = $duration
              SourcePath = $SourceFile; DestPath = $destPath; Skipped = $false }
}


# ═════════════════════════════════════════════════════════════════════════════
#  MAIN EXECUTION BLOCK
# ═════════════════════════════════════════════════════════════════════════════
try {
    # ── Load config ───────────────────────────────────────────────────────────
    if (-not (Test-Path $ConfigPath)) { throw "Config file not found: $ConfigPath" }
    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json

    # ── Resolve paths ─────────────────────────────────────────────────────────
    $sourceDir  = $config.ExportBaseDirectory   # e.g. E:\Datapump
    $targetDir  = if ($TargetBasePath -ne '') { $TargetBasePath } `
                  else { "\\$StageServer\f`$\DataImportFromProd" }

    # ── Open log file ─────────────────────────────────────────────────────────
    $logDir = $config.LogDirectory
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $Script:LogFile   = Join-Path $logDir "DataTransfer_${RunTimestamp}.log"
    $Script:LogWriter = [System.IO.StreamWriter]::new($Script:LogFile, $true, [System.Text.Encoding]::UTF8)
    $Script:LogWriter.AutoFlush = $true

    # ── Startup banner ────────────────────────────────────────────────────────
    Write-Log ('=' * 65)                                           -Level INFO -Phase 'STARTUP'
    Write-Log 'Oracle Stage Refresh - DMP File Transfer Pipeline'  -Level INFO -Phase 'STARTUP'
    Write-Log "Server (source) : $($env:COMPUTERNAME)"            -Level INFO -Phase 'STARTUP'
    Write-Log "Stage Server    : $StageServer"                     -Level INFO -Phase 'STARTUP'
    Write-Log "Source Dir      : $sourceDir"                       -Level INFO -Phase 'STARTUP'
    Write-Log "Target Dir      : $targetDir"                       -Level INFO -Phase 'STARTUP'
    Write-Log "Run ID          : $RunTimestamp"                    -Level INFO -Phase 'STARTUP'
    Write-Log "Log File        : $($Script:LogFile)"               -Level INFO -Phase 'STARTUP'
    Write-Log "Dry Run         : $($DryRun.IsPresent)"             -Level INFO -Phase 'STARTUP'
    Write-Log ('=' * 65)                                           -Level INFO -Phase 'STARTUP'

    # ── Verify source directory exists ────────────────────────────────────────
    if (-not (Test-Path $sourceDir)) { throw "Source directory not found: $sourceDir" }

    # ── Create target UNC directory if it does not exist ─────────────────────
    if (-not $DryRun.IsPresent) {
        if (-not (Test-Path $targetDir)) {
            Write-Log "Target directory not found - creating: $targetDir" -Level WARN -Phase 'STARTUP'
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            Write-Log "Target directory created: $targetDir" -Level SUCCESS -Phase 'STARTUP'
        }
        else {
            Write-Log "Target directory exists: $targetDir" -Level INFO -Phase 'STARTUP'
        }
    }

    # ── Verify UNC is reachable ───────────────────────────────────────────────
    if (-not $DryRun.IsPresent -and -not (Test-Path $targetDir)) {
        throw "Cannot reach target directory: $targetDir  (check network access and admin share on $StageServer)"
    }

    # ── Pipeline start email ──────────────────────────────────────────────────
    $schemaList   = ($config.Schemas | ForEach-Object { $_.Name }) -join ', '
    $allRecips    = @($config.Email.DBARecipients) + @($config.Email.BusinessRecipients)
    $dbaRecips    = @($config.Email.DBARecipients)
    $pipelineBody = @"
<h2>Oracle Stage Refresh &mdash; DMP Transfer Pipeline Started</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Source Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Stage Server</b></td><td>$StageServer</td></tr>
  <tr><td><b>Source Directory</b></td><td>$sourceDir</td></tr>
  <tr><td><b>Target Directory</b></td><td>$targetDir</td></tr>
  <tr><td><b>Schemas</b></td><td>$schemaList</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Dry Run</b></td><td>$($DryRun.IsPresent)</td></tr>
  <tr><td><b>Started</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
<p><i>Individual emails will follow for each schema file transfer.</i></p>
"@
    Send-TransferEmail -Subject "[TRANSFER STARTED] DMP Pipeline - Oracle Stage Refresh $RunTimestamp" `
                       -Body $pipelineBody -To $allRecips `
                       -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort

    # ── Process each schema ───────────────────────────────────────────────────
    $results = [System.Collections.ArrayList]@()

    foreach ($schemaConfig in $config.Schemas) {
        $schema = $schemaConfig.Name
        Write-Log "--- Schema: $schema ---" -Level INFO -Phase 'TRANSFER'

        # Find the most-recent .dmp file for this schema in the source directory
        $dmpFiles = @(Get-ChildItem -Path $sourceDir -Filter "${schema}_*.dmp" -File -ErrorAction SilentlyContinue |
                      Sort-Object LastWriteTime -Descending)

        if ($dmpFiles.Count -eq 0) {
            Write-Log "[$schema] No .dmp file found in $sourceDir - SKIPPING." -Level WARN -Phase 'TRANSFER'
            [void]$results.Add([PSCustomObject]@{
                Schema     = $schema; Success = $false; FileSizeMB = 0; Duration = 0
                SourceFile = 'NOT FOUND'; DestFile = '-'; Skipped = $false; Missing = $true
            })
            continue
        }

        $sourceFile = $dmpFiles[0].FullName
        Write-Log "[$schema] Latest dump: $sourceFile" -Level INFO -Phase 'TRANSFER'

        # Per-file start email
        $startBody = @"
<h3>DMP Transfer Started</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Source Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Stage Server</b></td><td>$StageServer</td></tr>
  <tr><td><b>Schema</b></td><td>$schema</td></tr>
  <tr><td><b>File</b></td><td>$(Split-Path $sourceFile -Leaf)</td></tr>
  <tr><td><b>Size</b></td><td>$([math]::Round((Get-Item $sourceFile).Length / 1MB,1)) MB</td></tr>
  <tr><td><b>Destination</b></td><td>$targetDir</td></tr>
  <tr><td><b>Started</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
"@
        Send-TransferEmail -Subject "[TRANSFER STARTED] $schema .dmp - Oracle Stage Refresh" `
                           -Body $startBody -To $allRecips `
                           -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort

        $result = Invoke-FileTransfer -Schema $schema -SourceFile $sourceFile `
                                      -TargetDir $targetDir -IsDryRun $DryRun.IsPresent

        [void]$results.Add([PSCustomObject]@{
            Schema     = $schema
            Success    = $result.Success
            FileSizeMB = $result.FileSizeMB
            Duration   = $result.Duration
            SourceFile = $result.SourcePath
            DestFile   = $result.DestPath
            Skipped    = $result.Skipped
            Missing    = $false
        })

        if ($result.Success) {
            $doneBody = @"
<h3 style='color:green'>DMP Transfer Completed Successfully</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Source Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Stage Server</b></td><td>$StageServer</td></tr>
  <tr><td><b>Schema</b></td><td>$schema</td></tr>
  <tr><td><b>File</b></td><td>$(Split-Path $result.SourcePath -Leaf)</td></tr>
  <tr><td><b>Destination</b></td><td>$($result.DestPath)</td></tr>
  <tr><td><b>Size Verified</b></td><td>$($result.FileSizeMB) MB</td></tr>
  <tr><td><b>Duration</b></td><td>$($result.Duration) min</td></tr>
  <tr><td><b>Completed</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
"@
            Send-TransferEmail -Subject "[TRANSFER COMPLETE] $schema - Oracle Stage Refresh" `
                               -Body $doneBody -To $allRecips `
                               -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort
        }
        else {
            $failBody = @"
<h3 style='color:red'>ALERT: DMP Transfer FAILED</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Source Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Stage Server</b></td><td>$StageServer</td></tr>
  <tr><td><b>Schema</b></td><td>$schema</td></tr>
  <tr><td><b>Source File</b></td><td>$($result.SourcePath)</td></tr>
  <tr><td><b>Destination</b></td><td>$($result.DestPath)</td></tr>
  <tr><td><b>Failed At</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
<p>Pipeline will continue with remaining schemas. Check attached log for details.</p>
"@
            $att = @(); if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $att += $Script:LogFile }
            Send-TransferEmail -Subject "[TRANSFER FAILED] $schema - Oracle Stage Refresh" `
                               -Body $failBody -To $dbaRecips `
                               -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort `
                               -Attachments $att
        }
    }

    # ── Final summary ─────────────────────────────────────────────────────────
    $successCount = @($results | Where-Object { $_.Success  }).Count
    $failCount    = @($results | Where-Object { -not $_.Success }).Count
    $totalSizeMB  = [math]::Round(($results | Measure-Object -Property FileSizeMB -Sum).Sum, 1)
    $totalMins    = [math]::Round(($results | Measure-Object -Property Duration   -Sum).Sum, 1)

    $tableRows = $results | ForEach-Object {
        $color  = if ($_.Success) { 'green' } else { 'red' }
        $status = if ($_.Skipped) { 'SKIPPED (DryRun)' } elseif ($_.Missing) { 'FILE NOT FOUND' } `
                  elseif ($_.Success) { 'SUCCESS' } else { 'FAILED' }
        "<tr><td>$($_.Schema)</td><td style='color:$color'><b>$status</b></td>" +
        "<td>$($_.FileSizeMB) MB</td><td>$($_.Duration) min</td><td>$(Split-Path $_.DestFile -Leaf)</td></tr>"
    }

    $overallStatus = if ($failCount -gt 0) { 'COMPLETED WITH ERRORS' } else { 'COMPLETED SUCCESSFULLY' }
    $summaryBody   = @"
<h2>Oracle Stage Refresh &mdash; DMP Transfer Summary</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Source Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Stage Server</b></td><td>$StageServer</td></tr>
  <tr><td><b>Target Directory</b></td><td>$targetDir</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Total Schemas</b></td><td>$($config.Schemas.Count)</td></tr>
  <tr><td><b>Successful</b></td><td style='color:green'><b>$successCount</b></td></tr>
  <tr><td><b>Failed</b></td><td style='color:red'><b>$failCount</b></td></tr>
  <tr><td><b>Total Data Transferred</b></td><td>${totalSizeMB} MB</td></tr>
  <tr><td><b>Total Duration</b></td><td>${totalMins} minutes</td></tr>
</table><br/>
<h3>Per-Schema Results</h3>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr style='background:#ddd'><th>Schema</th><th>Status</th><th>Size</th><th>Duration</th><th>Dest File</th></tr>
  $($tableRows -join '')
</table><br/><p>Full log attached.</p>
"@

    Write-Log "Transfer pipeline finished. Success=$successCount | Failed=$failCount | Total=${totalSizeMB} MB | ${totalMins} min" `
              -Level SUCCESS -Phase 'SUMMARY'

    # ── Copy TablespaceSnapshot.json to Stage ─────────────────────────────────
    $snapshotSrc  = Join-Path $sourceDir 'TablespaceSnapshot.json'
    $snapshotDest = Join-Path $targetDir 'TablespaceSnapshot.json'
    if (Test-Path $snapshotSrc) {
        if ($DryRun.IsPresent) {
            Write-Log "[SNAPSHOT] DRY RUN - would copy: $snapshotSrc -> $snapshotDest" -Level WARN -Phase 'SNAPSHOT'
        }
        else {
            try {
                Copy-Item -Path $snapshotSrc -Destination $snapshotDest -Force
                Write-Log "[SNAPSHOT] Copied TablespaceSnapshot.json to $snapshotDest" -Level SUCCESS -Phase 'SNAPSHOT'
            }
            catch {
                Write-Log "[SNAPSHOT] Failed to copy snapshot (non-fatal): $_" -Level WARN -Phase 'SNAPSHOT'
            }
        }
    }
    else {
        Write-Log "[SNAPSHOT] TablespaceSnapshot.json not found at $snapshotSrc — skipping." -Level WARN -Phase 'SNAPSHOT'
    }

    $summaryAtt = @(); if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $summaryAtt += $Script:LogFile }
    Send-TransferEmail -Subject "[TRANSFER SUMMARY] $overallStatus - Oracle Stage Refresh $RunTimestamp" `
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
<h2 style='color:red'>CRITICAL: DMP Transfer Pipeline Stopped</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Source Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Stage Server</b></td><td>$StageServer</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Error</b></td><td>$errMsg</td></tr>
  <tr><td><b>Time</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
</table>
<p>Please review the attached log immediately.</p>
"@
            $ca = @(); if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $ca += $Script:LogFile }
            Send-TransferEmail -Subject "[CRITICAL ERROR] DMP Transfer - $RunTimestamp" `
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
