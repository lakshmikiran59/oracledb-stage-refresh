@ECHO OFF
REM ====================================================================
REM DIS Schema Export - Standard Edition 2 Compatible
REM No Parallel, No Advanced Compression
REM ====================================================================

SETLOCAL EnableDelayedExpansion

SET ORACLE_HOME=E:\Oracle19CHome
SET PATH=%ORACLE_HOME%\bin;%PATH%
SET EXPORT_DIR=E:\Datapump

IF NOT EXIST "%EXPORT_DIR%" MKDIR "%EXPORT_DIR%"

ECHO ====================================================================
ECHO DIS Schema Export for AFIQA Refresh
ECHO Oracle Standard Edition 2 - Basic Mode
ECHO Starting: %DATE% %TIME%
ECHO ====================================================================
ECHO.

REM Create parameter file (NO PARALLEL, NO COMPRESSION)
(
ECHO USERID=oracledbas/orcldbadispatch64@dispatch
ECHO SCHEMAS=DIS
ECHO DIRECTORY=ORCL_DATAPUMP
ECHO DUMPFILE=DIS_Aprl1_2026_AFIQUARefresh.dmp
ECHO LOGFILE=DIS_Aprl1_2026_AFIQUARefresh.log
ECHO EXCLUDE=TABLE:"IN ('PACOS_LOCATIONHISTORY')"
ECHO CLUSTER=N
ECHO METRICS=YES
ECHO EXCLUDE=STATISTICS
) > "%EXPORT_DIR%\DIS_export.par"

ECHO Parameter File Contents:
ECHO ----------------------------------------
TYPE "%EXPORT_DIR%\DIS_export.par"
ECHO ----------------------------------------
ECHO.

ECHO Starting export...
ECHO WARNING: File will be uncompressed (larger size)
ECHO Ensure sufficient disk space is available
ECHO.

CD /D "%EXPORT_DIR%"
expdp parfile=DIS_export.par

IF %ERRORLEVEL% EQU 0 (
    ECHO.
    ECHO ====================================================================
    ECHO Export COMPLETED SUCCESSFULLY
    ECHO ====================================================================
    ECHO.
    DIR DIS_Aprl1_2026_AFIQUARefresh.dmp
) ELSE (
    ECHO.
    ECHO ====================================================================
    ECHO Export FAILED - Check log file
    ECHO ====================================================================
)

ECHO.
ECHO Log file: %EXPORT_DIR%\DIS_Aprl1_2026_AFIQUARefresh.log
ECHO.
PAUSE

ENDLOCAL
