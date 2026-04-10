# Pipeline Overview — Oracle Stage Refresh

## Purpose

This pipeline automates the periodic refresh of the Oracle Stage database (`aazeud-oracle03`) from Production (`aazeud-oracle02`). It replaces the manual bat file/SQL*Plus approach with a fully logged, email-notified PowerShell pipeline.

## Pipeline Flow

```mermaid
flowchart TD
    A([🚀 Start Refresh]) --> B

    subgraph S1["PHASE 1 — Export  |  aazeud-oracle02"]
        B["Start-OracleExport.ps1"] --> C["Init Oracle DIRECTORY\nORCL_DATAPUMP → E:\\Datapump"]
        C --> D["For each schema:\nGenerate .par file\nRun expdp"]
        D --> E["📁 E:\\Datapump\n SCHEMA_DATE_LABEL.dmp"]
    end

    subgraph S2["PHASE 2 — Transfer  |  aazeud-oracle02"]
        E --> F["Start-DataTransfer.ps1"]
        F --> G["Copy-Item UNC\nSource: E:\\Datapump\nTarget: \\\\StageServer\\f$\\DataImportFromProd"]
        G --> H["Size verification\nper file"]
    end

    subgraph S3["PHASE 3 — Import  |  aazeud-oracle03"]
        H --> I["Start-OracleImport.ps1"]
        I --> J["Init Oracle DIRECTORY\nORCL_DATAIMPORT → F:\\DataImportFromProd"]
        J --> K["For each schema:\nGenerate .par file\nRun impdp"]
        K --> L["Phase 3: PII Masking\n(if configured)"]
    end

    subgraph S4["PHASE 4 — Post-Import  |  aazeud-oracle03  |  MANUAL"]
        L --> M{{"📋 DBA Reviews\nImport Logs"}}
        M -->|Satisfied| N["Start-SequenceSyncAndStats.ps1"]
        N --> O["Step 1: Drop sequences"]
        O --> P["Step 2: Re-import sequences\nfrom prod dump"]
        P --> Q["Step 3: Gather Statistics\nDBMS_STATS.GATHER_SCHEMA_STATS"]
        Q --> R["Step 4: Recompile\nUTL_RECOMP.recomp_serial"]
    end

    R --> Z([✅ Refresh Complete])
```

## Data Flow

```mermaid
sequenceDiagram
    participant P as Production DB<br/>aazeud-oracle02
    participant D as E:\Datapump<br/>(Production)
    participant N as Network UNC
    participant S as F:\DataImportFromProd<br/>(Stage)
    participant Q as Stage DB<br/>aazeud-oracle03

    Note over P,D: Phase 1 — Export
    P->>D: expdp ASHLEY (SCHEMA_DATE_LABEL.dmp)
    P->>D: expdp CIM
    P->>D: expdp DIS (5 tables excluded)
    P->>D: expdp ... (5 more schemas)

    Note over D,S: Phase 2 — Transfer
    D->>N: Copy-Item (UNC)
    N->>S: ASHLEY_*.dmp (size verified)
    N->>S: CIM_*.dmp
    N->>S: DIS_*.dmp
    N->>S: ... (5 more schemas)

    Note over S,Q: Phase 3 — Import
    S->>Q: impdp ASHLEY (TABLE_EXISTS_ACTION=REPLACE)
    S->>Q: impdp CIM
    S->>Q: impdp DIS (1 index + 5 tables excluded)
    S->>Q: ... (5 more schemas)

    Note over Q: Phase 4 — Post-Import (Manual)
    Q->>Q: Drop + Re-import Sequences
    Q->>Q: Gather Statistics
    Q->>Q: Recompile Invalid Objects
```

## Server & Component Reference

| Component | Value |
|---|---|
| Production Server | `aazeud-oracle02` |
| Stage Server | `aazeud-oracle03` |
| Oracle Home (both) | `E:\Oracle19CHome` |
| Export Dump Directory | `E:\Datapump` |
| Export Logs | `E:\Datapump\Logs` |
| Stage Import Directory | `F:\DataImportFromProd` |
| Import Logs | `F:\DataImportFromProd\Logs` |
| Oracle Export DIR Object | `ORCL_DATAPUMP` |
| Oracle Import DIR Object | `ORCL_DATAIMPORT` |
| Dump File Pattern | `SCHEMA_MMMDD_YYYY_RUNLABEL.dmp` |

## Email Notifications

Every script sends HTML emails at each key event:

| Event | Recipients |
|---|---|
| Pipeline started | DBA + Business |
| Per-schema start | DBA + Business |
| Per-schema complete (with ORA- list) | DBA + Business |
| Per-schema failure (with log attached) | DBA only |
| Pipeline summary | DBA + Business |
| Critical error (pipeline stopped) | DBA only |
