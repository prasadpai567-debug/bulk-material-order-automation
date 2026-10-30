# Excel VBA SAP Purchase Order Automation

This project automates a controlled procurement workflow using Microsoft Excel VBA and SAP GUI Scripting.

The macro groups material-order rows by vendor, creates one purchase order per vendor in SAP transaction ME21N, and writes the extracted PO number to Excel column K.

> Safety warning: this macro can create real SAP purchase orders. Use synthetic data and a non-production SAP client before testing.

## Business problem

Material requests are often collected in Excel and then entered manually into SAP. This creates repetitive work and increases the risk of missing materials, incorrect quantities, incorrect dates, or incomplete purchasing information.

## Verified automation workflow

1. Read material-order rows from the active Excel worksheet.
2. Validate vendor, material, plant, quantity, delivery date, requisitioner, purchasing organization, and purchasing group.
3. Group rows by vendor.
4. Open SAP transaction ME21N.
5. Enter vendor and purchasing header data.
6. Add each vendor's material items.
7. Set delivery, invoice, and confirmation details.
8. Save the purchase order.
9. Extract the confirmed PO number from the SAP status bar.
10. Write the PO number to column K for the relevant Excel rows.

## Excel input columns

| Column | Field |
| --- | --- |
| A | Material |
| B | Vendor ID |
| C | Plant |
| D | PO Quantity |
| E | Delivery Date |
| F | Requisitioner |
| G | Purchasing Organization |
| H | Purchasing Group |
| K | Returned PO Number |

## Requirements

- Microsoft Excel desktop
- Macros enabled for the workbook
- SAP GUI for Windows
- SAP GUI Scripting enabled
- An active SAP session
- Permission to create purchase orders in ME21N
- Environment-specific SAP GUI element IDs

## How to run

1. Open the workbook in Excel.
2. Use synthetic test data.
3. Confirm that SAP is connected to a non-production client.
4. Verify that the worksheet columns match the documented input layout.
5. Run the SAP_Final_Solution macro or the workbook's RUN MACRO button.
6. Review SAP messages and the PO number written to column K.

## Error handling

The macro validates Excel data before connecting to SAP. It also checks SAP readiness, attempts to handle SAP warnings, verifies the SAP status bar after saving, and displays the current row, vendor, SAP item, and completed-PO count when an error occurs.

The current implementation does not create a persistent execution log or provide automatic rollback.

## Known limitations

- The macro uses environment-specific SAP GUI element IDs.
- The first SAP connection and session are selected automatically.
- The current worksheet is selected using ActiveSheet.
- Delivery time is fixed in the VBA source.
- Some notification and purchasing values are fixed in the VBA source.
- Dates are formatted using mm/dd/yyyy.
- Re-running the macro after an uncertain save may create duplicate purchase orders.
- No automated test suite or CI workflow is included.
- SAP, Excel, and VBA behavior cannot be verified in a headless environment.

## Security guidance

Do not commit SAP credentials, SAP connection files, supplier information, company exports, production screenshots, confidential workbook copies, or local file paths.

Use synthetic or fully redacted examples for demonstrations. Review the SAP client, session, GUI element IDs, fixed notification text, and purchasing values before operational use.

## Future improvements

- Add dry-run mode.
- Add explicit SAP client and session verification.
- Add duplicate-run protection.
- Add persistent logging.
- Add a synthetic test workbook.
- Add a controlled confirmation step before saving.
- Add screenshots or a demonstration using non-confidential data.

## Repository contents

- CRFR MATERIALS.xlsm - macro-enabled workbook and user interface.
- Module1.bas - exported VBA source for review and version control.
- README.md - project documentation.

## Portfolio summary

This project demonstrates Excel VBA development, SAP GUI Scripting, procurement workflow automation, input validation, transaction integration, and operational risk awareness.
