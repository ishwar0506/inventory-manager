# Inventory Manager

A desktop-based inventory, stock-transfer, and order-management
application built with **Python, Tkinter, and SQLite**. It provides
separate sections for inward stock, outward stock, current stock
summaries, stock transfers, and customer orders. Data is stored locally
in an SQLite database.

## Features

### 1. Inward Inventory

-   Record incoming stock with entry date, memo number, merchant, item
    name, and quantity.
-   Search by item name and filter by merchant or memo number.
-   Optionally display records for a selected date.
-   Import inward records from a CSV file.
-   Export inward records to a formatted Excel workbook.

### 2. Outward Inventory

-   Record outgoing stock with entry date, bale number, transport, item
    name, and quantity.
-   Search by item name and filter by transport or bale number.
-   Automatically remember the most recently used bale number in
    `last_bale_no.txt`; the `*` entry shortcut can advance a trailing
    bale number.
-   Look up and export the contents of a specific bale number.
-   Import outward records from CSV and export them to Excel.

### 3. Stock Summary

-   Calculate net stock for each item as:

    `Net Stock = Total Inward Quantity - Total Outward Quantity`

-   Search by item and optionally filter the summary to a selected date.

-   Export the summary to CSV or Excel.

-   Start a stock transfer from a selected item.

### 4. Stock Transfers

-   View transfer history, filter by date, and search by item.
-   Inspect transfers involving an item by double-clicking a transfer
    row.
-   Export an individual item's transfer history or all transfer records
    to Excel.
-   View basic transfer statistics, including total transfers and unique
    source and destination items.

### 5. Orders

-   Record orders using party name, agent, quality name, date, and
    quantity.
-   Search orders by party and optionally filter by date.
-   View quantities grouped by party/agent and quality, with row and
    overall totals.
-   Export order data to Excel.

### 6. Additional functionality

-   Autocomplete suggestions for relevant text-entry fields.
-   Quantity unit conversion: entering a quantity with the suffix `d`
    converts dozens to pieces (`1d = 12 pcs`). Plain numeric quantities
    are treated as pieces.
-   Keyboard shortcuts: `Ctrl+Z` for undo and `Ctrl+Y` for redo.
-   `F2` focuses the date field in the active tab.
-   Right-click context menus and keyboard actions are available for
    selected inventory and transfer records.

> **Note:** Undo/redo support depends on the action. The code explicitly
> reports that delete operations cannot be undone. Treat undo/redo as a
> convenience, not as a substitute for backups.

## Technology Stack

-   **Python** --- application logic
-   **Tkinter / ttk** --- desktop graphical user interface
-   **SQLite (`sqlite3`)** --- local database
-   **OpenPyXL** --- Excel workbook import/export support
-   **tkcalendar** --- date-picker widgets
-   **CSV** --- CSV import and summary export

## Requirements

-   Python 3
-   Tkinter (usually included with standard Python installations; on
    some Linux distributions it must be installed separately)
-   `openpyxl`
-   `tkcalendar`

## Installation

### 1. Get the project

Place `king_manager.py` in a folder where you want the application
and its database to be stored.

### 2. Install dependencies

On Ubuntu/Debian:

``` bash
sudo apt update
sudo apt install python3-tk
python3 -m pip install openpyxl tkcalendar
```

If your system uses a virtual environment, activate it before installing
the Python packages.

### 3. Run the application

``` bash
python3 "king_manager.py"
```

The application window is titled **Inventory Manager**.

## Data and Files

The application creates and uses these local files in its working
directory:

  -----------------------------------------------------------------------
  File                                Purpose
  ----------------------------------- -----------------------------------
  `basic_inventory.db`                SQLite database containing inward
                                      inventory, outward inventory,
                                      stock-transfer, and order records

  `last_bale_no.txt`                  Stores the last used bale number
                                      for the outward-entry workflow
  -----------------------------------------------------------------------

The database tables are created automatically when the application
starts. The program also checks for certain columns and adds them when
they are missing.

**Important:** Keep `basic_inventory.db` safe. It contains the
application's records. Close the application before manually copying the
database for a backup.

## Basic Usage

1.  Launch the program.
2.  Select the required tab: **Inward**, **Outward**, **Summary**,
    **Transfer**, or **Orders**.
3.  Enter the relevant details and quantity.
4.  Use the search fields and date filters to narrow the displayed
    records.
5.  Use the import/export controls to move data between the application
    and CSV/Excel files.
6.  Use the Summary tab to review net stock and initiate transfers.

### Quantity format

-   `25` --- 25 pieces
-   `2d` --- 24 pieces
-   `-1d` --- -12 pieces

The only unit suffix defined in the current code is `d` (dozens). Other
suffixes are rejected.

### CSV import format

For Inward and Outward imports, the expected column order is:

``` csv
Date,Item Name,Quantity,Memo No,Merchant
14/06/25,Rice,50,M001,ABC Store
14/06/25,Wheat,30,M002,XYZ Supplier
```

For Outward imports, use `Bale No` and `Transport` as the fourth and
fifth columns instead:

``` csv
Date,Item Name,Quantity,Bale No,Transport
14/06/25,Rice,20,B001,ABC Transport
```

The first row may contain headers. Supported date formats include
`DD/MM/YY`, `YYYY-MM-DD`, `DD-MM-YYYY`, `MM/DD/YYYY`, and `DD/MM/YYYY`.
Empty rows are skipped. If an imported record matches an existing
item/reference/date/entity combination, its quantity is added to the
existing quantity.

## Keyboard Shortcuts

  -----------------------------------------------------------------------
  Shortcut                            Action
  ----------------------------------- -----------------------------------
  `Ctrl + Z`                          Undo the last supported action

  `Ctrl + Y`                          Redo the last undone supported
                                      action

  `F2`                                Focus the date field in the active
                                      tab

  `Delete`                            Delete the selected inventory item
                                      or order/quality entry where
                                      supported

  `Double-click`                      Edit an inventory item name; in the
                                      Summary tab, initiate a stock
                                      transfer; in Transfer history,
                                      inspect transfers for an item
  -----------------------------------------------------------------------

Some actions are tab-specific. Context-menu actions may also be
available with a right-click.

## Notes and Limitations

-   This is a local desktop application; the code does not implement a
    server, cloud sync, user accounts, or multi-user access.
-   The database is stored relative to the program's current working
    directory. Run the application from the intended project folder to
    keep its data files together.
-   CSV import is implemented for Inward and Outward inventory. Orders
    and stock transfers have Excel export functionality.
-   The current quantity conversion map defines only `d` for dozens.
-   Deleting records is destructive; the undo manager does not restore
    deleted records.
-   Make regular backups of `basic_inventory.db`, especially before
    importing large files or making bulk changes.

## Project Structure

``` text
project-folder/
├── king_manager.py
├── basic_inventory.db       # created automatically on first run
└── last_bale_no.txt         # created/updated by the outward-entry workflow
```

## License

No license is specified in the provided source file. Add a license
before distributing or publishing the project if you want to define
reuse and redistribution terms.
