# Ortems Toolbox v3.0 - Change Log and Review Notes

## Scope
This version refactors the uploaded AutoIt application that creates DELMIA Ortems demonstration data directly in SQL Server. The Excel workbook `Toolbox v2.0.1.xlsm` was reviewed as the functional rule model, especially the hidden MODULES sheet that documents cross-table relationships.

## Main UI updates
- Renamed the application window to **Ortems Toolbox v3.0**.
- Reworked the top title/subtitle and bottom action bar for clearer workflow.
- Renamed the global Excel actions to **Import Workbook...** and **Export Workbook...**.
- Added tooltips to clarify that the Excel workflow is now workbook-based, not folder/CSV-based.
- Adjusted bottom button resizing so the Integrity Check button stays aligned when the window is resized.
- Improved row creation/edit dialogs with clearer save/cancel behavior and validation before a row is saved.

## Data-entry validation added
The generic row dialog now trims values and blocks saving when common formats are wrong:
- Day fields must be numeric values from 1 to 7.
- Time fields must use `HH:MM`.
- Date fields must use `dd/mm/yyyy` or `dd/mm/yyyy hh:mm` depending on the field.
- Numeric fields such as quantity, duration, setup, break, resource count, and phase must be numeric.
- Core required fields are checked for Calendars, Machines, Operations, Routings, Items, and Work Orders.
- Machines with CT Type `4` are allowed without a Calendar ID.
- Raw material items (`MP`) are allowed without a Routing ID.

## Excel import/export improvements
The previous global import/export behavior was folder/CSV based. It has been replaced by a round-trip Excel workbook flow:
- Export creates one `.xlsx` workbook.
- Workbook includes a README sheet plus one sheet per data tab.
- Supported sheets: Calendars, Machines, Operations, Routings, Items, BOM, WorkOrders, WOLinks, SecondaryResources, Capacity, InventoryMovements.
- Headers are exported in row 3 and data starts in row 4.
- All exported cells are formatted as text to reduce Excel date/time/ID conversion issues.
- Import validates sheet names and headers before replacing current data.
- Import regenerates the Line column inside the application.
- After import, the Integrity Check runs automatically and reports issues.

## Integrity Check improvements
- Added required-field checks for core datasets.
- Added duplicate-key checks for critical keys:
  - Calendar + Start Day + Start Time
  - Machine ID
  - WC ID + Machine ID
  - Operation ID + WC ID + Machine ID
  - Routing ID + Phase Code
  - Item ID + Version
  - Work Order ID
  - WO Link composite relationship
  - Capacity calendar + Start Day + Start Time
- Kept BOM validation intentionally focused on Routing ID only, matching the requested behavior that BOM integrity should ignore all columns except Routing ID.
- Kept the STAND-BY machine exception.
- Kept the CT Type `4` calendar exception.
- Updated item/routing integrity so non-raw-material items must reference an existing routing, while MP/raw-material items remain exempt.

## Notes and limitation
- The file was reviewed and patched statically in this environment.
- AutoIt execution and Microsoft Excel COM automation require Windows with AutoIt and Excel installed, so the application could not be compiled or run end-to-end here.
- The code uses Excel COM for workbook import/export; this means Microsoft Excel must be installed on the target Windows machine.
- All comments and user-facing text added or changed in this version are in English.
