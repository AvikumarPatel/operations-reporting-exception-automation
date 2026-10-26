# Setup guide

## Prerequisites

- Windows with Microsoft Excel desktop
- Microsoft 365 organizational account
- Power Automate Desktop
- Power Automate Cloud access
- OneDrive for Business
- Office Scripts access in Excel
- Outlook and optionally Microsoft Teams

## Workbook structure

Create a macro-enabled workbook with these sheets:

- `Raw_Data`
- `KPI_Summary`
- `Exceptions`
- `Config`

### Config values

| Setting | Example value |
|---|---:|
| ErrorRateThreshold | 0.02 |
| LateOrderRateThreshold | 0.05 |
| MinimumOrdersPerHour | 15 |
| ReturnRateThreshold | 0.03 |

## Local / desktop stage

1. Import `src/vba/ProcessOperationsReport.bas` into a standard VBA module.
2. Load sample or test data into `Raw_Data`.
3. Run the macro manually once to confirm workbook behavior.
4. Build the PAD flow documented in `flows/power-automate-desktop.md`.
5. Confirm PAD can open the workbook, run the macro, read KPI values, save, and close Excel.

## Cloud stage

1. Save a macro-free `.xlsx` copy after local processing.
2. Upload it to OneDrive for Business.
3. Create the Office Script from `src/office-scripts/ReviewOperationsKPIs.ts`.
4. Test the Office Script directly from Excel Online.
5. Create the manual and/or scheduled cloud flows from `flows/power-automate-cloud.md`.
6. Configure Outlook and Teams connections with an organizational Microsoft 365 account.

## Important

Do not commit real customer data, production credentials, tenant secrets, exported connection references, or personal access tokens to a public repository.
