# Metal Inventory — iPhone PWA prototype

## What works
- Import an existing `.xlsx` file with an `Inventory_Master` worksheet (uses SheetJS via CDN, internet required for initial library load).
- View dashboard counts and cost estimates, browse/search/filter inventory, edit **non-formula input columns**, save pending changes locally, export changes as CSV, and back up local data as JSON.
- Reimport an updated workbook; blocks import on detected conflicting edits.
- App uses browser localStorage; no inventory is uploaded to a server by this prototype.

## What is NOT implemented
- **No automatic or live two-way synchronization.** CSV edits must be reviewed/applied in Excel separately. This prototype intentionally does not write `.xlsx` files because doing so can damage complex formulas, formatting, connections, and workbook features.
- No multi-device synchronization, user authentication, live spot-price API, or guaranteed offline support.
- No create/delete inventory records yet. No formula recalculation. Dashboard numbers are illustrative derived summaries, not equivalent to Excel's dashboards.

## How to run
1. Extract this ZIP to a folder.
2. On a Windows PC, serve the folder with `python -m http.server 8000` and open `http://localhost:8000` in Edge. Or deploy the folder to an HTTPS static hosting service.
3. On iPhone, open the HTTPS URL in Safari and use Share > Add to Home Screen. iPhone installation requires an accessible HTTPS URL; opening the ZIP itself does not install an app.
4. Open **Excel Sync**, import your workbook, then browse inventory and edit records.
5. Export pending edits CSV and review it before applying changes to your original Excel workbook.

## Recommended next architecture for genuine two-way sync
- Keep the original workbook backed up and authoritative until migration is validated.
- Use a server-side database (record IDs, revision numbers, change history) with authenticated HTTPS API.
- Add an Excel connector: a controlled Power Query data import + a separately authenticated Excel-to-server write path (e.g. Office Script/Power Automate or dedicated add-in), not an unsafe rewrite of the original workbook.
- Conflict detection should be per field with revision numbers and user-confirmed resolutions.
- Decide whether Excel is stored in OneDrive/SharePoint or a local directory; this determines feasible sync integration.


## Sync controls (November 2026 update)

The app now displays **Sync Now**, **auto-sync on app open**, and **auto-sync after saving edits**. Preferences are retained locally. These are **not yet connected** to OneDrive: the Sync Now action explicitly reports that no data was transmitted and preserves pending edits. This is intentional to prevent accidental workbook corruption or false success messages.

The selected test workbook is `Metal_Inventory_iPhone_TEST.xlsx`, not the master. The OneDrive browser URL is an identifier, **not** an API access token.

### Requirements before activating sync
1. Register a public-client or SPA Microsoft identity application supporting personal Microsoft accounts; obtain client ID and configure exact hosted redirect URI. Never embed a client secret in browser code.
2. Use authorization-code with PKCE, with minimal delegated file permissions and a securely hosted HTTPS origin.
3. Build a protected service that reads and writes the **entire XLSX file** through Microsoft Graph personal OneDrive file APIs (not Excel workbook editing APIs). Store no passwords.
4. Read and verify the test file's drive item ID, ETag, file contents, and inventory schema. Make a recoverable backup before writes.
5. For writes, re-fetch and check ETag/version, merge only approved nonformula fields by stable InventoryID, preserve all other workbook content and relationships, reject conflicts, and revalidate the resulting XLSX before upload. Use a conditional upload and retry/abort on conflict.
6. Test both directions and formula/Power Query preservation on a copy before ever considering the master.

**Current safety state:** No OAuth sign-in, Graph requests, cloud storage, or background synchronization is implemented. No automatic synchronization can run while the iPhone app is closed. iOS may suspend PWAs in the background.
