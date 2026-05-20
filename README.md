# Google-Sheets-Warehouse-Pick-List-Generator
Automate your warehouse fulfillment with this Google Apps Script. It extracts "ReadyToShip" orders, cross-references inventory tables to locate product locations and images, and builds a colorful, printable pick list. Features include a custom UI popup with WhatsApp integration for quickly dispatching instructions to your warehouse team.

# Google Sheets Warehouse Pick List Generator

An automated script to simplify warehouse order fulfillment directly within Google Sheets.

## Features
* **Automated Extraction:** Pulls "ReadyToShip" orders from your workflow and maps them to actual inventory locations and images.
* **Beautiful Formatting:** Automatically generates a "PickList" tab formatted with professional headers, alternating row colors, and inline product images.
* **Interactive Dashboard:** Opens an interactive HTML popup displaying picking stats (order lines, quantities, locations, and missing item warnings).
* **WhatsApp Integration:** Generates formatted dispatch messages and allows you to instantly ping your warehouse team via WhatsApp web links.

## Setup Instructions

1. **Access Apps Script:** Open your Google Sheet, click `Extensions` > `Apps Script`.
2. **Add the Script:** Paste the sanitized code into the editor.
3. **Configure Your IDs:** At the top of the file, locate these two variables and replace the placeholder text with your actual spreadsheet IDs (found in their URLs):
   * `PICK_LIST_ORDER_WORKFLOW_SPREADSHEET_ID`
   * `PICK_LIST_INVENTORY_SPREADSHEET_ID`
4. **Save:** Save the file (`Ctrl/Cmd + S`).
5. **Run:** Select the `generatePickList` function from the dropdown menu in the editor toolbar and click `Run` (authorize the script if prompted), or link the `generatePickList` function to a custom drawing/button in your sheet.
