# Architecture

```mermaid
flowchart TD
    A[Operational data in Excel] --> B[Excel VBA: ProcessOperationsReport]
    B --> C[Clean + validate records]
    C --> D[Calculate KPIs]
    D --> E[Flag operational and data-quality exceptions]
    E --> F[KPI_Summary + Exceptions sheets]
    F --> G[Power Automate Desktop]
    G --> H[Run VBA + read KPI cells]
    H --> I{ExceptionCount > 0?}
    I -->|Yes| J[Desktop exception alert]
    I -->|No| K[Continue]
    J --> L[Save and close workbook]
    K --> L
    L --> M[Processed .xlsx in OneDrive for Business]
    M --> N[Office Script: ReviewOperationsKPIs]
    N --> O[Return structured KPI object]
    O --> P{Cloud exceptionCount > 0?}
    P -->|True| Q[Outlook review-required email]
    P -->|True| R[Teams notification]
    P -->|False| S[Outlook no-exceptions email]
```

## Separation of responsibilities

- **VBA:** local data preparation, validation, KPI calculations, exception detection, and worksheet output.
- **Power Automate Desktop:** desktop orchestration around Excel and VBA.
- **Office Scripts:** cloud-friendly KPI extraction, formatting, timestamping, and structured return values.
- **Power Automate Cloud:** scheduling, branching, Outlook notifications, and Teams routing.
- **OneDrive for Business:** cloud storage for the processed workbook used by Excel Online (Business).
