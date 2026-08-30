# RIS Portal

Google Apps Script source for the Requisition Issuance Slip Portal.

This version reads inventory directly from two source tabs instead of using the old consolidated inventory sheet:

- Medicine: `DRUGS and MEDICINES`
- Supplies: `2026 SUPPLIES`

The old `RIS Consolidated Inventory` tab and its menu can be left in the workbook while testing this version, but this portal does not read from it.

## Files

- `Code.gs` - portal backend, sheet setup, RIS submission, source routing, stock checks, and email notifications.
- `Index.html` - web app frontend with Medicine/Supplies buttons and stock-status UI.
- `RIS_Excel_Generator_Addon.gs` - generates an Excel RIS file from the configured RIS Excel template.
- `RIS_PDF_Generator_Addon.gs` - generates a legal-size RIS PDF.
- `RIS_Status_Menu.gs` - sends request and delivery status emails.
- `appsscript.json` - Apps Script manifest with required scopes and Drive advanced service entry.

## Required Configuration

Update these placeholder IDs before using the system:

In `Code.gs`:

```js
transactionsSheetId: 'PASTE_TRANSACTIONS_SHEET_ID',
inventorySheetId: 'PASTE_INVENTORY_SHEET_ID',
```

In `RIS_Excel_Generator_Addon.gs`:

```js
templateFileId: 'PASTE_RIS_EXCEL_TEMPLATE_FILE_ID',
outputFolderId: 'PASTE_EXCEL_OUTPUT_FOLDER_ID',
```

In `RIS_PDF_Generator_Addon.gs`:

```js
outputFolderId: 'PASTE_PDF_OUTPUT_FOLDER_ID',
```

If the Apps Script project is bound to the transaction Google Sheet, `transactionsSheetId` can stay as the placeholder while testing. The inventory workbook ID is still recommended because the portal reads Medicine and Supplies from the inventory workbook.

## Required Google Sheets

Run `RIS Portal > Setup required sheets` after pasting the files into Apps Script. It creates or updates:

- `RIS Sources`
- `RIS Entries`
- `RIS Items`
- `Emails`
- `Holidays`
- `Users`

## RIS Sources

The `RIS Sources` sheet controls the buttons shown in the portal.

Default rows:

| Enabled | Button Name | Source Spreadsheet ID | Source Sheet Name | Header Row | Category |
| --- | --- | --- | --- | --- | --- |
| TRUE | Medicine | same | DRUGS and MEDICINES | auto | Medicine |
| TRUE | Supplies | same | 2026 SUPPLIES | auto | Supplies |

Use `same` only when the inventory tabs are in the same spreadsheet as the Apps Script project or when `inventorySheetId` is configured in `Code.gs`.

## Inventory Headers

The code supports flexible header names. The uploaded CHD Logistics workbook uses:

Medicine tab `DRUGS and MEDICINES`:

- `ITEM CODE`
- `ITEM DESCRIPTION`
- `UNIT OF MEASURE`
- `BATCH / LOT NUMBER`
- `EXPIRATION DATE`
- `SUPPLIER`
- ` PURCHASE ORDER NUMBER`
- `DELIVEY DATE`
- `END USER`
- `PHYSICAL COUNT`
- `UNIT COST`
- `TOTAL COST`
- `REMARKS`

Supplies tab `2026 SUPPLIES`:

- `ITEM CODE`
- `DESCRIPTION`
- `UNIT OF MEASURE`
- `BATCH / LOT NUMBER`
- `EXPIRATION DATE (MMDDYYYY)`
- `SUPPLIER`
- `PURCHASE ORDER NUMBER`
- `DELIVERY DATE`
- `END USER`
- `PHYSICAL COUNT`
- `UNIT COST`
- `TOTAL AMOUNT`
- `REMARKS`

The typo `DELIVEY DATE`, the leading space in ` PURCHASE ORDER NUMBER`, and the difference between `ITEM DESCRIPTION` and `DESCRIPTION` are handled by header aliases.

## RIS Excel Generator

The Excel generator is based on the uploaded `RIS BLANK FORMAT.xlsx` template:

- Sheet: first sheet by default
- Item rows: `10:40`
- Item columns: `B:N`
- Total formula: `M41 = SUM(M10:M40)`
- Purpose cell: `C42`

Before filling the Excel file, the generator validates item details from `RIS Items`. If a value is blank, it attempts to recover it from the source inventory row using `Source Sheet` and `Source Row`. This is intended to prevent the old issue where Excel fields were blank even though the inventory/database had values.

Required item fields for Excel generation:

- Item Code
- Item Description
- UOM
- Qty Requested
- Issued Qty
- Unit Cost

Recoverable fields:

- Item Code
- Item Description
- UOM
- PO/PTR Number
- Supplier
- Batch / Lot No.
- Expiry Date
- Unit Cost
- Remarks

## Required Apps Script Services

Enable the Advanced Drive service for Excel template conversion and Excel export:

1. Open Apps Script.
2. Click **Services**.
3. Add **Drive API**.

The manifest already includes the Drive advanced service entry, but it may still need to be enabled in the Apps Script editor.

## Authorization

Run these once from the spreadsheet menus:

- `RIS Portal > Authorize email notifications`
- `RIS Excel Tools > Authorize Excel generator`
- `RIS PDF Tools > Authorize PDF generator`

The Excel and PDF tools validate users against the `Users` sheet. Required headers:

```text
First Name | Last Name | Username | User Type | Status
```

Only users with `Status` equal to `Active` can generate Excel/PDF outputs.

## Deployment

1. Create or open the Google Sheet that will store RIS transactions.
2. Open **Extensions > Apps Script**.
3. Add the `.gs` and `.html` files from this repository.
4. Update configuration IDs.
5. Enable the Advanced Drive service.
6. Run `setupRISPortalSheets`.
7. Run each authorization menu item once.
8. Deploy as a web app:
   - Execute as: Me
   - Who has access: choose based on your office workflow

## Testing Checklist

- Medicine button loads programs from `DRUGS and MEDICINES`.
- Supplies button loads programs from `2026 SUPPLIES`.
- Out-of-stock rows are red and cannot be added.
- Low-stock rows are amber.
- Submitted RIS writes one row to `RIS Entries`.
- Submitted RIS writes line items to `RIS Items`.
- `TESTING_MODE = true` does not deduct stock.
- `TESTING_MODE = false` deducts stock from the source inventory row.
- Excel generation fills item code, description, UOM, PO/supplier, batch, expiry, quantity, cost, and total.
- PDF generation writes PDF URL back to `RIS Entries`.
- Excel generation writes Excel URL back to `RIS Entries`.
- Status menu updates `RIS Status` and sends the correct email.
