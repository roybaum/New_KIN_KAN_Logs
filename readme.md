# New KIN KAN Logs

Google Apps Script project for the KIN/KAN log workbook, including Inventory sync tooling.

## Project Purpose

- Maintain KIN/KAN workbook helpers and menu actions.
- Keep the workbook `Inventory` tab current.
- Support webhook-based ingest for external inventory updates.

## Key Files

- `src/index.ts`: Main Apps Script source (menus, inventory import sync, navigation, triggers).
- `WOAFR_ingest_script.gs`: Web App endpoint (`doPost`) for webhook ingest into `Inventory`.
- `Invoke-MasterInventorySync.ps1`: Scheduled PowerShell wrapper used on the WideOrbit central server.

## Inventory Sync Runbook

This section documents the production flow where a scheduled PowerShell job exports inventory from WideOrbit and posts batches to the Apps Script webhook.

### 1. Data Flow

1. Windows Task Scheduler runs `Invoke-MasterInventorySync.ps1` on the WideOrbit central server.
2. The script resolves settings from parameters and environment variables.
3. The script invokes the exporter (`Export-MediaAssetsToMasterInventory.ps1`) to collect rows.
4. Exported rows are sent to the Apps Script Web App endpoint.
5. `WOAFR_ingest_script.gs` validates the token and writes to the workbook `Inventory` sheet.

### 2. Required Environment Variables

- `WO_SERVER_IP`: WideOrbit host/IP.
- `WO_MASTER_WEBHOOK_URL`: Deployed Apps Script Web App URL.
- `WO_MASTER_WEBHOOK_TOKEN`: Shared secret token configured in Script Properties.

### 3. Optional Environment Variables

- `WO_MASTER_SHEET_NAME`: Target sheet name (defaults to `Inventory`).
- `WO_MASTER_WRITE_MODE`: `upsert` or `replace` (defaults to `upsert`).
- `WO_CATEGORIES`: Comma-delimited 3-letter categories (for example: `COM,PSA,PRM`).

### 4. Example Manual Run

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-MasterInventorySync.ps1 -Verbose
```

Optional explicit categories:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-MasterInventorySync.ps1 -Categories COM,PSA
```

### 5. Task Scheduler Suggested Action

- Program/script: `powershell.exe`
- Arguments:

```text
-NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\Invoke-MasterInventorySync.ps1"
```

- Start in: Folder that contains the script.

### 6. Logging and Verification

- Script transcript logs are written to `Logs\MasterInventorySync-YYYYMMDD-HHMMSS.log`.
- Verify successful webhook responses and processed row counts in the latest log.
- In Apps Script execution logs, confirm no token/auth errors.
- In the workbook, confirm the `Inventory` tab was updated.

### 7. Troubleshooting

- Error: `Missing required configuration`:
   - Set `WO_SERVER_IP`, `WO_MASTER_WEBHOOK_URL`, and `WO_MASTER_WEBHOOK_TOKEN`.
- Error: `Unauthorized request token`:
   - Token in server env var does not match Apps Script `MASTER_WEBHOOK_TOKEN`.
- Error: `Could not find exporter script`:
   - Ensure `Export-MediaAssetsToMasterInventory.ps1` exists at the path expected by `Invoke-MasterInventorySync.ps1`.
- No rows updated:
   - Confirm categories contain data and verify exporter output before webhook submission.

### 8. Security Notes

- Do not commit real webhook URLs with sensitive query params or real tokens.
- Keep secrets in environment variables or secure secret storage.
- Rotate `WO_MASTER_WEBHOOK_TOKEN` if server access changes.

## Development Notes

- This repo is TypeScript-first (`src/index.ts`) with generated output in `dist/index.js`.
- `WOAFR_ingest_script.gs` may be deployed separately as a Web App endpoint.