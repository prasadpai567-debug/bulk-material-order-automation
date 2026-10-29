# bulk-material-order-automation

This repository contains an Excel/VBA automation that groups material-order rows by vendor, creates one SAP purchase order per vendor through SAP GUI Scripting and writes the returned PO number to Excel column K.

## Contents

- `CRFR MATERIALS.xlsm` â€” macro-enabled workbook and user interface.
- `Module1.bas` â€” exported VBA source for review and version control.
- Optional walkthrough video â€” include only when publishing with `-IncludeVideo`.

## Runtime requirements

- Microsoft Excel desktop with macros enabled for this trusted workbook.
- SAP GUI for Windows with SAP GUI Scripting enabled and an active SAP session.
- Access to transaction `ME21N` and permission to create purchase orders.

## Expected input columns

The macro reads the active worksheet: A Material, B Vendor ID, C Plant, D PO Quantity, E Delivery Date, F Requisitioner, G Purchasing Organization and H Purchasing Group. It writes the created PO number to column K.

## Important safety notes

Run only with test data in a non-production SAP client first. The macro creates real SAP purchase orders when connected to a live session. Do not commit passwords, SAP connection files, tokens or confidential exports. Review the fixed notification text and SAP GUI element IDs for your environment before use.
