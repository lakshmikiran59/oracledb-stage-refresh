<#
.SYNOPSIS
    Oracle Stage Refresh — Standalone Sequence Sync + Statistics Gather.

.DESCRIPTION
    Run this script AFTER you have reviewed the import logs and are satisfied
    the data import is acceptable (even if some errors were ignored).

    For each selected schema the script performs three steps:
      Step 1 - DROP   : Drops all existing sequences in the schema on Stage.
      Step 2 - IMPORT : Re-imports sequences from the production dump (INCLUDE=SEQUENCE).
      Step 3 - STATS  : Runs DBMS_STATS.GATHER_SCHEMA_STATS (CASCADE=TRUE, degree=1).

    Any step can be skipped independently via switch parameters.

.PARAMETER Schemas
    Comma-separated list of schema names to process.
    Defaults to ALL schemas defined in ImportConfig.json.
    Example: -Schemas ASHLEY,CIM,DIS

.PARAMETER ConfigPath
    Path to ImportConfig.json. Defaults to .\Config\ImportConfig.json.

.PARAMETER SkipDrop
    Skip Step 1 — do NOT drop existing sequences before importing.
    Use when you want to ADD missing sequences without removing current ones.

.PARAMETER SkipSequenceSync
    Skip Steps 1 and 2 entirely — only gather statistics.

.PARAMETER SkipStats
    Skip Step 3 — do NOT gather statistics after sequence sync.

.PARAMETER DryRun
    Validates config and generates par files but does NOT execute any
    SQL or impdp commands.

.EXAMPLE
    # Full run for all schemas
    .\Start-SequenceSyncAndStats.ps1

    # Specific schemas only
    .\Start-SequenceSyncAndStats.ps1 -Schemas ASHLEY,CIM,DIS

    # Skip drop — only import sequences + gather stats
    .\Start-SequenceSyncAndStats.ps1 -SkipDrop

    # Only gather statistics (no sequence work)
    .\Start-SequenceSyncAndStats.ps1 -SkipSequenceSync

    # Sequence sync only, no stats
    .\Start-SequenceSyncAndStats.ps1 -SkipStats

    # Dry run — safe to test
    .\Start-SequenceSyncAndStats.ps1 -DryRun

.NOTES
    Compatible : PowerShell 5.0+
    Oracle Ed. : Standard Edition 2 (degree=1, CLUSTER=N)
    Runs on    : Stage server (aazeud-oracle03)
    Requires   : ImportConfig.json in .\Config\
    Author     : DBA Team - Ashley Furniture India
#>

[CmdletBinding()]
param(
    [string[]]$Schemas         = @(),
    [string]  $ConfigPath      = "$PSScriptRoot\Config\ImportConfig.json",
    [switch]  $SkipDrop,
    [switch]  $SkipSequenceSync,
    [switch]  $SkipStats,
    [switch]  $SkipRecompile,
    [switch]  $DryRun
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
        [string]$Phase = 'SEQSYNC'
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
# FUNCTION: Send-Email
# ─────────────────────────────────────────────────────────────────────────────
function Send-Email {
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
# FUNCTION: Invoke-DropSequences   [STEP 1]
# Drops all existing sequences in the schema on Stage.
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-DropSequences {
    param([string]$Schema, [string]$SqlPlusExe, [bool]$IsDryRun)

    Write-Log "[$Schema] Step 1 - Dropping existing sequences on Stage..." -Level INFO -Phase 'DROP'

    if ($IsDryRun) {
        Write-Log "[$Schema] DRY RUN - drop skipped." -Level WARN -Phase 'DROP'
        return $true
    }

    $sql = @"
SET SERVEROUTPUT ON SIZE UNLIMITED
SET FEEDBACK OFF
WHENEVER SQLERROR CONTINUE

DECLARE
    v_count NUMBER := 0;
BEGIN
    FOR s IN (SELECT sequence_name
              FROM   dba_sequences
              WHERE  sequence_owner = UPPER('$Schema')
              ORDER  BY sequence_name) LOOP
        BEGIN
            EXECUTE IMMEDIATE 'DROP SEQUENCE $Schema.' || s.sequence_name;
            v_count := v_count + 1;
            DBMS_OUTPUT.PUT_LINE('DROPPED : $Schema.' || s.sequence_name);
        EXCEPTION
            WHEN OTHERS THEN
                DBMS_OUTPUT.PUT_LINE('WARN    : Could not drop ' || s.sequence_name || ' - ' || SQLERRM);
        END;
    END LOOP;
    DBMS_OUTPUT.PUT_LINE('---');
    DBMS_OUTPUT.PUT_LINE('TOTAL   : Dropped ' || v_count || ' sequences from $Schema');
END;
/
EXIT;
"@

    $result = Invoke-SqlPlus -SqlBlock $sql -SqlPlusExe $SqlPlusExe
    foreach ($line in $result.Output) {
        if    ($line -match 'WARN|ORA-|ERROR') { Write-Log "  $line" -Level WARN  -Phase 'DROP' }
        elseif ($line.Trim() -ne '')            { Write-Log "  $line" -Level INFO  -Phase 'DROP' }
    }
    if ($result.ExitCode -ne 0) {
        Write-Log "[$Schema] Drop sequences failed (exit $($result.ExitCode))." -Level ERROR -Phase 'DROP'
        return $false
    }
    Write-Log "[$Schema] Sequence drop completed." -Level SUCCESS -Phase 'DROP'
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-ImportSequences   [STEP 2]
# Re-imports sequences from the production dump using INCLUDE=SEQUENCE.
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-ImportSequences {
    param(
        [PSCustomObject]$SchemaConfig,
        [string]        $ImportBaseDir,
        [string]        $OracleDirectoryObject,
        [string]        $DumpFileName,
        [string]        $OracleHome,
        [bool]          $IsDryRun
    )
    $schema     = $SchemaConfig.Name
    $schemaDir  = Join-Path $ImportBaseDir $schema
    if (-not (Test-Path $schemaDir)) { New-Item -ItemType Directory -Path $schemaDir -Force | Out-Null }

    $seqLogFile = "${schema}_seqimport_${RunTimestamp}.log"
    $seqParPath = Join-Path $schemaDir "${schema}_seqsync.par"

    $lines = @(
        "USERID=""$($SchemaConfig.Credential)""",
        "SCHEMAS=$schema",
        "DIRECTORY=$OracleDirectoryObject",
        "DUMPFILE=$DumpFileName",
        "LOGFILE=$seqLogFile",
        "INCLUDE=SEQUENCE",
        "CLUSTER=N",
        "METRICS=YES"
    )
    $lines | Set-Content -Path $seqParPath -Encoding ASCII
    Write-Log "[$schema] Step 2 - Par file: $seqParPath" -Level INFO -Phase 'SEQIMPORT'
    Write-Log "[$schema] Step 2 - Importing sequences from prod dump: $DumpFileName" -Level INFO -Phase 'SEQIMPORT'

    if ($IsDryRun) {
        Write-Log "[$schema] DRY RUN - impdp sequence import skipped." -Level WARN -Phase 'SEQIMPORT'
        return $true
    }

    $impdpExe = Join-Path $OracleHome 'bin\impdp.exe'
    $proc     = Start-Process -FilePath $impdpExe `
                              -ArgumentList "parfile=`"$seqParPath`"" `
                              -WorkingDirectory $schemaDir `
                              -NoNewWindow -Wait -PassThru
    $exitCode = $proc.ExitCode

    # Log any ORA- errors from impdp sequence log
    $oraLog = Join-Path $schemaDir $seqLogFile
    if (Test-Path $oraLog) {
        $errors = @(Select-String -Path $oraLog -Pattern 'ORA-' | Select-Object -First 20)
        $errors | ForEach-Object { Write-Log "  $($_.Line)" -Level WARN -Phase 'SEQIMPORT' }
    }

    if ($exitCode -ne 0) {
        Write-Log "[$schema] Sequence import failed (exit $exitCode)." -Level ERROR -Phase 'SEQIMPORT'
        return $false
    }
    Write-Log "[$schema] Sequence import completed successfully." -Level SUCCESS -Phase 'SEQIMPORT'
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-GatherStatistics   [STEP 3]
# Gathers schema statistics. degree=1 for SE2 (no parallel).
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-GatherStatistics {
    param([string]$Schema, [string]$SqlPlusExe, [bool]$IsDryRun)

    Write-Log "[$Schema] Step 3 - Gathering schema statistics..." -Level INFO -Phase 'STATS'

    if ($IsDryRun) {
        Write-Log "[$Schema] DRY RUN - statistics gather skipped." -Level WARN -Phase 'STATS'
        return $true
    }

    $sql = @"
SET SERVEROUTPUT ON SIZE UNLIMITED
SET FEEDBACK OFF
WHENEVER SQLERROR EXIT SQL.SQLCODE

BEGIN
    DBMS_OUTPUT.PUT_LINE('STATS: Start - $Schema - ' || TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS'));
    DBMS_STATS.GATHER_SCHEMA_STATS(
        ownname => UPPER('$Schema'),
        cascade => TRUE,
        options => 'GATHER',
        degree  => 1
    );
    DBMS_OUTPUT.PUT_LINE('STATS: Done  - $Schema - ' || TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS'));
END;
/
EXIT;
"@

    $result = Invoke-SqlPlus -SqlBlock $sql -SqlPlusExe $SqlPlusExe
    foreach ($line in $result.Output) {
        if ($line -match 'ORA-|ERROR') { Write-Log "  $line" -Level ERROR -Phase 'STATS' }
        elseif ($line.Trim() -ne '')   { Write-Log "  $line" -Level INFO  -Phase 'STATS' }
    }
    if ($result.ExitCode -ne 0) {
        Write-Log "[$Schema] Statistics gather FAILED (exit $($result.ExitCode))." -Level ERROR -Phase 'STATS'
        return $false
    }
    Write-Log "[$Schema] Statistics gather COMPLETED." -Level SUCCESS -Phase 'STATS'
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
# FUNCTION: Invoke-RecompileInvalidObjects   [STEP 4]
# Recompiles all invalid objects in the schema using UTL_RECOMP.recomp_serial.
# Reports invalid object counts before and after. degree=1, SE2 safe.
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-RecompileInvalidObjects {
    param([string]$Schema, [string]$SqlPlusExe, [bool]$IsDryRun)

    Write-Log "[$Schema] Step 4 - Recompiling invalid objects..." -Level INFO -Phase 'RECOMP'

    if ($IsDryRun) {
        Write-Log "[$Schema] DRY RUN - recompile skipped." -Level WARN -Phase 'RECOMP'
        return $true
    }

    $sql = @"
SET SERVEROUTPUT ON SIZE UNLIMITED
SET FEEDBACK OFF
WHENEVER SQLERROR EXIT SQL.SQLCODE

DECLARE
    v_before NUMBER;
    v_after  NUMBER;
BEGIN
    -- Count invalid objects before recompilation
    SELECT COUNT(*) INTO v_before
    FROM   dba_objects
    WHERE  owner  = UPPER('$Schema')
    AND    status = 'INVALID';

    DBMS_OUTPUT.PUT_LINE('RECOMP: Invalid objects before : ' || v_before);

    -- Recompile all invalid objects serially (SE2 safe - no parallel)
    UTL_RECOMP.recomp_serial(UPPER('$Schema'));

    -- Count remaining invalid objects after recompilation
    SELECT COUNT(*) INTO v_after
    FROM   dba_objects
    WHERE  owner  = UPPER('$Schema')
    AND    status = 'INVALID';

    DBMS_OUTPUT.PUT_LINE('RECOMP: Invalid objects after  : ' || v_after);
    DBMS_OUTPUT.PUT_LINE('RECOMP: Successfully fixed     : ' || (v_before - v_after));

    IF v_after > 0 THEN
        DBMS_OUTPUT.PUT_LINE('RECOMP: WARNING - ' || v_after || ' object(s) still invalid after recompile.');
    ELSE
        DBMS_OUTPUT.PUT_LINE('RECOMP: All objects compiled successfully.');
    END IF;
END;
/
EXIT;
"@

    $result = Invoke-SqlPlus -SqlBlock $sql -SqlPlusExe $SqlPlusExe
    foreach ($line in $result.Output) {
        if ($line -match 'ORA-|ERROR')  { Write-Log "  $line" -Level ERROR -Phase 'RECOMP' }
        elseif ($line -match 'WARNING') { Write-Log "  $line" -Level WARN  -Phase 'RECOMP' }
        elseif ($line.Trim() -ne '')    { Write-Log "  $line" -Level INFO  -Phase 'RECOMP' }
    }
    if ($result.ExitCode -ne 0) {
        Write-Log "[$Schema] Recompile FAILED (exit $($result.ExitCode))." -Level ERROR -Phase 'RECOMP'
        return $false
    }
    Write-Log "[$Schema] Recompile COMPLETED." -Level SUCCESS -Phase 'RECOMP'
    return $true
}

# ═════════════════════════════════════════════════════════════════════════════
#  MAIN EXECUTION BLOCK
# ═════════════════════════════════════════════════════════════════════════════
try {
    # ── Load config ───────────────────────────────────────────────────────────
    if (-not (Test-Path $ConfigPath)) { throw "Config file not found: $ConfigPath" }
    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json

    $env:ORACLE_HOME = $config.OracleHome
    $env:PATH        = "$($config.OracleHome)\bin;$($env:PATH)"
    $sqlplusExe      = Join-Path $config.OracleHome 'bin\sqlplus.exe'
    if (-not (Test-Path $sqlplusExe)) { throw "sqlplus.exe not found: $sqlplusExe" }

    # ── Open log ──────────────────────────────────────────────────────────────
    $logDir = $config.LogDirectory
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $Script:LogFile   = Join-Path $logDir "SeqSyncStats_${RunTimestamp}.log"
    $Script:LogWriter = [System.IO.StreamWriter]::new($Script:LogFile, $true, [System.Text.Encoding]::UTF8)
    $Script:LogWriter.AutoFlush = $true

    # ── Filter schemas ────────────────────────────────────────────────────────
    $schemasToRun = if ($Schemas.Count -gt 0) {
        $upper = $Schemas | ForEach-Object { $_.ToUpper().Trim() }
        @($config.Schemas | Where-Object { $upper -contains $_.Name.ToUpper() })
    } else { @($config.Schemas) }

    if ($schemasToRun.Count -eq 0) { throw "No matching schemas found. Check -Schemas parameter." }

    # ── Startup banner ────────────────────────────────────────────────────────
    Write-Log ('=' * 65)                                                              -Level INFO -Phase 'STARTUP'
    Write-Log 'Oracle Stage Refresh - Sequence Sync + Statistics'                    -Level INFO -Phase 'STARTUP'
    Write-Log "Server         : $($env:COMPUTERNAME)"                                -Level INFO -Phase 'STARTUP'
    Write-Log "Run ID         : $RunTimestamp"                                        -Level INFO -Phase 'STARTUP'
    Write-Log "Schemas        : $(($schemasToRun | ForEach-Object {$_.Name}) -join ', ')" -Level INFO -Phase 'STARTUP'
    Write-Log "Skip Drop      : $($SkipDrop.IsPresent)"                              -Level INFO -Phase 'STARTUP'
    Write-Log "Skip Seq Sync  : $($SkipSequenceSync.IsPresent)"                      -Level INFO -Phase 'STARTUP'
    Write-Log "Skip Stats     : $($SkipStats.IsPresent)"                             -Level INFO -Phase 'STARTUP'
    Write-Log "Skip Recompile : $($SkipRecompile.IsPresent)"                         -Level INFO -Phase 'STARTUP'
    Write-Log "Dry Run        : $($DryRun.IsPresent)"                                -Level INFO -Phase 'STARTUP'
    Write-Log ('=' * 65)                                                              -Level INFO -Phase 'STARTUP'

    $allRecips = @($config.Email.DBARecipients) + @($config.Email.BusinessRecipients)
    $results   = [System.Collections.ArrayList]@()

    # ── Process each schema ───────────────────────────────────────────────────
    foreach ($schemaConfig in $schemasToRun) {
        $schema    = $schemaConfig.Name
        $dropOk      = $true
        $seqOk       = $true
        $statsOk     = $true
        $recompileOk = $true

        Write-Log ('─' * 65) -Level INFO -Phase 'STARTUP'
        Write-Log "Processing schema: $schema" -Level INFO -Phase 'STARTUP'
        Write-Log ('─' * 65) -Level INFO -Phase 'STARTUP'

        if (-not $SkipSequenceSync.IsPresent) {
            # Find most-recent dump for this schema
            $dmpFiles = @(Get-ChildItem -Path $config.ImportBaseDirectory -Filter "${schema}_*.dmp" -File -ErrorAction SilentlyContinue |
                          Sort-Object LastWriteTime -Descending)

            if ($dmpFiles.Count -eq 0) {
                Write-Log "[$schema] No .dmp file found in $($config.ImportBaseDirectory) - skipping sequence sync." -Level WARN -Phase 'SEQIMPORT'
                $dropOk = $false; $seqOk = $false
            }
            else {
                $dumpFileName = $dmpFiles[0].Name
                Write-Log "[$schema] Using dump: $dumpFileName" -Level INFO -Phase 'SEQIMPORT'

                # Step 1 - Drop sequences
                if (-not $SkipDrop.IsPresent) {
                    $dropOk = Invoke-DropSequences -Schema $schema -SqlPlusExe $sqlplusExe -IsDryRun $DryRun.IsPresent
                } else {
                    Write-Log "[$schema] Step 1 - SKIPPED (-SkipDrop)" -Level WARN -Phase 'DROP'
                }

                # Step 2 - Import sequences
                $seqOk = Invoke-ImportSequences -SchemaConfig          $schemaConfig `
                                                 -ImportBaseDir         $config.ImportBaseDirectory `
                                                 -OracleDirectoryObject $config.OracleDirectoryObject `
                                                 -DumpFileName          $dumpFileName `
                                                 -OracleHome            $config.OracleHome `
                                                 -IsDryRun              $DryRun.IsPresent
            }
        }
        else {
            Write-Log "[$schema] Steps 1 & 2 - SKIPPED (-SkipSequenceSync)" -Level WARN -Phase 'SEQIMPORT'
        }

        # Step 3 - Gather statistics
        if (-not $SkipStats.IsPresent) {
            $statsOk = Invoke-GatherStatistics -Schema $schema -SqlPlusExe $sqlplusExe -IsDryRun $DryRun.IsPresent
        }
        else {
            Write-Log "[$schema] Step 3 - SKIPPED (-SkipStats)" -Level WARN -Phase 'STATS'
        }

        # Step 4 - Recompile invalid objects
        if (-not $SkipRecompile.IsPresent) {
            $recompileOk = Invoke-RecompileInvalidObjects -Schema $schema -SqlPlusExe $sqlplusExe -IsDryRun $DryRun.IsPresent
        }
        else {
            Write-Log "[$schema] Step 4 - SKIPPED (-SkipRecompile)" -Level WARN -Phase 'RECOMP'
        }

        [void]$results.Add([PSCustomObject]@{
            Schema       = $schema
            DropOk       = $dropOk
            SeqOk        = $seqOk
            StatsOk      = $statsOk
            RecompileOk  = $recompileOk
        })
    }

    # ── Summary ───────────────────────────────────────────────────────────────
    Write-Log ('=' * 65) -Level INFO -Phase 'SUMMARY'
    Write-Log 'SEQUENCE SYNC + STATISTICS - SUMMARY' -Level INFO -Phase 'SUMMARY'
    Write-Log ('=' * 65) -Level INFO -Phase 'SUMMARY'

    $tableRows = $results | ForEach-Object {
        $dropStatus  = if ($SkipDrop.IsPresent -or $SkipSequenceSync.IsPresent) { 'SKIPPED' }
                       elseif ($_.DropOk)  { '✔ OK' } else { '✘ FAILED' }
        $seqStatus   = if ($SkipSequenceSync.IsPresent) { 'SKIPPED' }
                       elseif ($_.SeqOk)   { '✔ OK' } else { '✘ FAILED' }
        $statsStatus    = if ($SkipStats.IsPresent)     { 'SKIPPED' } elseif ($_.StatsOk)     { '✔ OK' } else { '✘ FAILED' }
        $recompStatus   = if ($SkipRecompile.IsPresent) { 'SKIPPED' } elseif ($_.RecompileOk) { '✔ OK' } else { '✘ FAILED' }

        $dropColor    = if ($dropStatus   -eq '✔ OK') { 'green' } elseif ($dropStatus   -eq 'SKIPPED') { 'gray' } else { 'red' }
        $seqColor     = if ($seqStatus    -eq '✔ OK') { 'green' } elseif ($seqStatus    -eq 'SKIPPED') { 'gray' } else { 'red' }
        $statsColor   = if ($statsStatus  -eq '✔ OK') { 'green' } elseif ($statsStatus  -eq 'SKIPPED') { 'gray' } else { 'red' }
        $recompColor  = if ($recompStatus -eq '✔ OK') { 'green' } elseif ($recompStatus -eq 'SKIPPED') { 'gray' } else { 'red' }

        Write-Log "$($_.Schema.PadRight(12)) | Drop: $dropStatus | Seq: $seqStatus | Stats: $statsStatus | Recompile: $recompStatus" -Level INFO -Phase 'SUMMARY'

        "<tr><td>$($_.Schema)</td>" +
        "<td style='color:$dropColor'><b>$dropStatus</b></td>" +
        "<td style='color:$seqColor'><b>$seqStatus</b></td>" +
        "<td style='color:$statsColor'><b>$statsStatus</b></td>" +
        "<td style='color:$recompColor'><b>$recompStatus</b></td></tr>"
    }

    $summaryBody = @"
<h2>Oracle Stage Refresh &mdash; Sequence Sync + Statistics Summary</h2>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr><td><b>Server</b></td><td>$($env:COMPUTERNAME)</td></tr>
  <tr><td><b>Run ID</b></td><td>$RunTimestamp</td></tr>
  <tr><td><b>Schemas</b></td><td>$(($results | ForEach-Object {$_.Schema}) -join ', ')</td></tr>
  <tr><td><b>Skip Drop</b></td><td>$($SkipDrop.IsPresent)</td></tr>
  <tr><td><b>Skip Seq Sync</b></td><td>$($SkipSequenceSync.IsPresent)</td></tr>
  <tr><td><b>Skip Stats</b></td><td>$($SkipStats.IsPresent)</td></tr>
  <tr><td><b>Skip Recompile</b></td><td>$($SkipRecompile.IsPresent)</td></tr>
</table><br/>
<table border='1' cellpadding='5' style='border-collapse:collapse;font-family:Arial'>
  <tr style='background:#ddd'><th>Schema</th><th>Step 1 - Drop Seq</th><th>Step 2 - Seq Import</th><th>Step 3 - Statistics</th><th>Step 4 - Recompile</th></tr>
  $($tableRows -join '')
</table><br/><p>Full log attached.</p>
"@
    $summaryAtt = @(); if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $summaryAtt += $Script:LogFile }
    Send-Email -Subject "[SEQ SYNC + STATS SUMMARY] Oracle Stage Refresh $RunTimestamp" `
               -Body $summaryBody -To $allRecips `
               -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort `
               -Attachments $summaryAtt
}
catch {
    $errMsg = $_.Exception.Message
    Write-Log "CRITICAL: $errMsg" -Level ERROR -Phase 'FATAL'
    try {
        if ($null -ne $config -and $null -ne $config.Email) {
            $ca = @(); if ($Script:LogFile -and (Test-Path $Script:LogFile)) { $ca += $Script:LogFile }
            Send-Email -Subject "[CRITICAL ERROR] Seq Sync + Stats - $RunTimestamp" `
                       -Body "<h3 style='color:red'>Error: $errMsg</h3><p>Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>" `
                       -To @($config.Email.DBARecipients) `
                       -From $config.Email.From -SMTPServer $config.Email.SMTPServer -SMTPPort $config.Email.SMTPPort `
                       -Attachments $ca
        }
    } catch { Write-Host "Failed to send error email: $_" -ForegroundColor Red }
    exit 1
}
finally {
    if ($null -ne $Script:LogWriter) {
        try   { $Script:LogWriter.Close(); $Script:LogWriter.Dispose() }
        catch { Write-Host "Warning: could not close log writer: $_" -ForegroundColor Yellow }
        $Script:LogWriter = $null
    }
}

