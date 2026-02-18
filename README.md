# 📊 Excel XLL Add-ins Collection

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A collection of **ready-to-use `.xll` add-ins** that extend Microsoft Excel with powerful features not available by default. Each add-in is built with **C# and Excel-DNA**, runs natively inside Excel, and requires **no macros or VBA**.

> **Works with:** Excel 2010, 2013, 2016, 2019, and Microsoft 365 (Desktop)

---

## 📥 Quick Start — How to Install Any Add-in

### Step 1: Check Your Excel Version (32-bit or 64-bit)

You must install the XLL that matches your Excel's architecture.

1. Open **Excel**.
2. Click **File → Account → About Excel**.
3. At the top of the popup, it will say either **32-bit** or **64-bit**.

> [!TIP]
> Most modern installations are **64-bit**. If in doubt, try the 64-bit version first.

### Step 2: Download the Correct XLL File

Each add-in folder below contains an **`Addin/`** subfolder with two files:
- `*_64bit.xll` — for **64-bit** Excel (most common)
- `*_32bit.xll` — for **32-bit** Excel

### Step 3: Install the Add-in

1. Open Excel → **File → Options → Add-ins**.
2. At the bottom, set **Manage** to **Excel Add-ins** and click **Go…**
3. Click **Browse…** and select the downloaded `.xll` file.
4. If prompted, click **Yes** to copy the add-in to your library folder.
5. Ensure it is **checked** in the list and click **OK**.

> [!IMPORTANT]
> If you see a security warning, right-click the `.xll` file → **Properties** → check **Unblock** at the bottom → click **OK**, then try again.

### Step 4: Verify

The add-in is now loaded. Its features (ribbon tab, formulas, etc.) will be available immediately.

---

## 📦 Add-ins Included

| # | Add-in | Description | Type |
|---|--------|-------------|------|
| 1 | [XLOOKUP for Older Excel](#1--xlookup-for-older-excel) | Brings the `XLOOKUP` function to Excel 2010–2019 | Formula (UDF) |
| 2 | [Excel Crosshair Highlighter](#2--excel-crosshair-highlighter) | Highlights the active row and column as you navigate | Ribbon Toggle |
| 3 | [Strikethrough Filter](#3--strikethrough-filter) | Filters rows based on strikethrough formatting | Ribbon Button |
| 4 | [Excel Delimiter Text Functions](#4--excel-delimiter-text-functions) | Delimiter-aware `LEFT`, `RIGHT`, and `MID` functions | Formula (UDF) |
| 5 | [Excel Filter Copy](#5--excel-filter-copy) | Copies visible (filtered) cells with full formatting | Ribbon Button |
| 6 | [Excel Merged Filter](#6--excel-merged-filter) | Filters data in sheets with merged cells | Ribbon Button |
| 7 | [Excel Navigator Arrows](#7--excel-navigator-arrows) | Focus-mode row-by-row navigation for filtered data | Ribbon Toolbar |
| 8 | [Replace Many](#8--replace-many) | Bulk find-and-replace using a mapping table | Formula + Ribbon |

---

## 1 — XLOOKUP for Older Excel

**Folder:** [`XLOOKUP-Excel-2019-2016-2013-2010/`](./XLOOKUP-Excel-2019-2016-2013-2010/)

Brings the powerful `XLOOKUP` function (normally only in Office 365) to **Excel 2010, 2013, 2016, and 2019**.

### Usage

```excel
=XLOOKUP(lookup_value, lookup_array, return_array)
```

Works exactly like the native Office 365 `XLOOKUP`, including:
- Exact match by default
- Customizable `match_mode` and `search_mode`
- Returns `#N/A` when no match is found (or your custom `if_not_found` value)

### Files

| File | Description |
|------|-------------|
| `Addin/XLookup_64bit.xll` | Add-in for 64-bit Excel |
| `Addin/XLookup_32bit.xll` | Add-in for 32-bit Excel |
| `Source/XLookupFunctions.cs` | Core XLOOKUP implementation |
| `Source/XLookupAddIn.csproj` | Project file |

---

## 2 — Excel Crosshair Highlighter

**Folder:** [`Excel-Crosshair-Highlighter/`](./Excel-Crosshair-Highlighter/)

Highlights the **entire active row and column** with a soft color as you navigate your worksheet. Gives you a visual crosshair that follows your cursor.

### Usage

1. After installation, a **Crosshair** tab appears in the Excel Ribbon.
2. Click the toggle button to enable/disable the highlighter.
3. Press **any arrow key** once after enabling — the highlight will then follow smoothly.

### Features

- Viewport-aware highlighting (only highlights visible area)
- Cached scope for flicker-free performance
- Persistent across sessions — loads automatically with Excel
- Works in all open workbooks simultaneously

### Files

| File | Description |
|------|-------------|
| `Addin/ExcelCrosshairAddIn_64bit.xll` | Add-in for 64-bit Excel |
| `Addin/ExcelCrosshairAddIn_32bit.xll` | Add-in for 32-bit Excel |
| `Source/CrosshairAddIn.cs` | Add-in entry point |
| `Source/CrosshairController.cs` | Core crosshair logic |
| `Source/RibbonController.cs` | Ribbon UI integration |
| `Source/ExcelCrosshairAddIn.csproj` | Project file |

---

## 3 — Strikethrough Filter

**Folder:** [`Strikethrough-Filter-Excel-AddIn/`](./Strikethrough-Filter-Excel-AddIn/)

Filters rows based on **strikethrough formatting** in a selected column. Useful when you use strikethrough to mark completed or cancelled items and want to hide them.

### Usage

1. Select any cell in the column you want to filter by.
2. Use the **Strikethrough Filter** button in the Ribbon.
3. Enter the header row number when prompted.
4. Rows with strikethrough text will be hidden; run again to restore.

### Features

- Toggle-safe: run again to fully restore the sheet
- Supports partial strikethrough in cells
- Uses Excel's built-in AutoFilter
- Never deletes user data — only hides/shows rows
- Optimized for up to ~50,000 rows

### Files

| File | Description |
|------|-------------|
| `Addin/StrikethroughFilterAddIn_64bit.xll` | Add-in for 64-bit Excel |
| `Addin/StrikethroughFilterAddIn_32bit.xll` | Add-in for 32-bit Excel |
| `Source/StrikethroughController.cs` | Core filter logic |
| `Source/StrikethroughFilterAddIn.csproj` | Project file |

---

## 4 — Excel Delimiter Text Functions

**Folder:** [`Excel-Delimiter-Text-Functions/`](./Excel-Delimiter-Text-Functions/)

Provides delimiter-aware alternatives to Excel's `LEFT`, `RIGHT`, and `MID` functions. Extract text segments based on a delimiter character instead of a character count.

### Usage

```excel
=TextLeft("A - B - C - D", "-", 2)       → "A - B"
=TextRight("A - B - C - D", "-", 3)      → "D"
=TextMid("ISO - 25A1 - 12345 - P3", "-", 2, 3) → "12345"
```

### Functions

| Function | Description |
|----------|-------------|
| `TextLeft(text, delimiter, n)` | Returns text **before** the Nth delimiter |
| `TextRight(text, delimiter, n)` | Returns text **after** the Nth-from-last delimiter |
| `TextMid(text, delimiter, n1, n2)` | Returns text **between** the N1th and N2th delimiters |

### Files

| File | Description |
|------|-------------|
| `Addin/TextDelimiterAddIn_64bit.xll` | Add-in for 64-bit Excel |
| `Addin/TextDelimiterAddIn_32bit.xll` | Add-in for 32-bit Excel |
| `Source/TextDelimiterFunctions.cs` | UDF implementations |
| `Source/TextDelimiterController.cs` | Ribbon UI integration |
| `Source/TextDelimiterForm.cs` | Interactive form UI |
| `Source/TextDelimiterAddIn.csproj` | Project file |

---

## 5 — Excel Filter Copy

**Folder:** [`Excel-Filter-Copy/`](./Excel-Filter-Copy/)

Copies visible (filtered) cells to a destination column while **preserving all formatting**, including merged cells. Excel's native copy-paste on filtered data often breaks merged cell structure — this add-in solves that.

### Usage

1. Apply a filter on your data.
2. Select the source range (the filtered column).
3. Click the **Copy Filtered** button in the Ribbon.
4. Click any cell in the destination column when prompted.
5. Data is pasted with full formatting and merge structure preserved.

### Features

- Handles merged cells correctly during copy
- Preserves all cell formatting (colors, fonts, borders, etc.)
- Performance-optimized with screen updating disabled during operation
- Clear error messages if destination is protected

### Files

| File | Description |
|------|-------------|
| `Addin/ExcelFilterCopy_64bit.xll` | Add-in for 64-bit Excel |
| `Addin/ExcelFilterCopy_32bit.xll` | Add-in for 32-bit Excel |
| `Source/AddIn.cs` | Add-in entry point and ribbon handler |
| `Source/CopyEngine.cs` | Core copy logic with merge handling |
| `Source/ExcelFilterCopy.csproj` | Project file |
| `Source/Properties/AssemblyInfo.cs` | Assembly metadata |

---

## 6 — Excel Merged Filter

**Folder:** [`Excel-Merged-Filter-AddIn/`](./Excel-Merged-Filter-AddIn/)

Enables **filtering on sheets with merged cells** — something Excel cannot do natively. When you try to filter a column with merged cells, Excel either throws an error or produces incorrect results. This add-in handles it properly.

### Usage

1. Click the **Merged Filter** button in the Ribbon.
2. A form appears — select the column and filter criteria.
3. The add-in applies the filter while respecting merged cell boundaries.
4. Use the **Clear Filter** button to restore the original view.

### Features

- Filter dropdown UI built specifically for merged cell data
- Preserves merged cell structure while filtering
- State management to track and restore original view
- Handles complex multi-row merged blocks

### Files

| File | Description |
|------|-------------|
| `Addin/ExcelMergedFilter_64bit.xll` | Add-in for 64-bit Excel |
| `Addin/ExcelMergedFilter_32bit.xll` | Add-in for 32-bit Excel |
| `Source/AddIn.cs` | Add-in entry point |
| `Source/FilterEngine.cs` | Core filtering logic |
| `Source/FilterForm.cs` | Filter selection dialog |
| `Source/FilterForm.Designer.cs` | Form designer code |
| `Source/ExcelUtils.cs` | Excel helper utilities |
| `Source/StateManager.cs` | Filter state tracking |
| `Source/ExcelMergedFilter.csproj` | Project file |
| `Source/Properties/AssemblyInfo.cs` | Assembly metadata |

---

## 7 — Excel Navigator Arrows

**Folder:** [`Excel-Navigator-Arrows/`](./Excel-Navigator-Arrows/)

A **focus-mode navigator** for filtered data. Navigate filtered rows one at a time using arrow buttons, with all other rows hidden for a clean, distraction-free view.

### Usage

1. Apply a filter on your data.
2. Click **Start Focus** in the Navigator Arrows ribbon tab.
3. Use the **← Previous** and **Next →** buttons to move through visible rows one at a time.
4. Click **Show List** to see all filtered rows in a summary sheet.
5. Click **Stop Focus** to restore the normal view.

### Features

- Single-row focus mode — shows only one filtered row at a time
- Previous/Next navigation with status bar indicator
- "Show List" creates a summary sheet of all filtered entries
- Double-click any row in the list to jump to it
- Workbook protection to prevent accidental sheet deletion during focus mode
- Crash recovery — automatically unprotects workbooks on restart
- Safe for 20,000+ rows

### Files

| File | Description |
|------|-------------|
| `Addin/NavigatorArrowsAddIn_64bit.xll` | Add-in for 64-bit Excel |
| `Addin/NavigatorArrowsAddIn_32bit.xll` | Add-in for 32-bit Excel |
| `Source/NavigatorController.cs` | Core navigation and UI logic |
| `Source/AddinLifecycle.cs` | Startup/shutdown and crash recovery |
| `Source/NavigatorArrowsAddIn.csproj` | Project file |

---

## 8 — Replace Many

**Folder:** [`Replace-Many-Excel/`](./Replace-Many-Excel/)

A high-performance **bulk find-and-replace** tool using a mapping table. Performs full-word matching to prevent accidental partial replacements (e.g., replaces "Cat" but not "Category").

### Usage — Formula (UDF)

```excel
=REPLACE_MANY(data_range, mapping_range)
```

The mapping range should be a two-column table with "From" and "To" values.

### Usage — Popup Macro

1. Click the **Replace Many** button in the Ribbon.
2. Select the mapping range (From → To table).
3. Choose the scope: **Selection**, **Active Sheet**, or **Entire Workbook**.
4. Choose case sensitivity.
5. Click **Replace** — all matching text is replaced in-place.

### Features

- **Full-word matching** — prevents partial replacements
- **Length priority** — longer keys are matched first to prevent clobbering
- **Array-based processing** — works in memory for high performance
- **Flexible scope** — apply to selection, sheet, or entire workbook
- **Case-sensitive option** — toggle case sensitivity as needed
- Duplicate keys: first occurrence wins; blank keys are ignored

### Mapping Format Example

| From | To |
|------|-----|
| USD | US Dollar |
| EUR | Euro |
| GBP | British Pound |

### Files

| File | Description |
|------|-------------|
| `Addin/ReplaceManyAddIn_64bit.xll` | Add-in for 64-bit Excel |
| `Addin/ReplaceManyAddIn_32bit.xll` | Add-in for 32-bit Excel |
| `Source/ReplaceManyFunctions.cs` | UDF implementation |
| `Source/ReplaceManyController.cs` | Ribbon UI and popup logic |
| `Source/ReplaceManyOptionsForm.cs` | Options form for scope/case settings |
| `Source/ReplaceManyAddIn.csproj` | Project file |

---

## 🏗️ Building from Source

All add-ins are built with **C# (.NET Framework 4.7.2)** and **[Excel-DNA](https://excel-dna.net/)**.

### Prerequisites

- Visual Studio 2019 or later (or `dotnet build`)
- .NET Framework 4.7.2 SDK
- NuGet packages will restore automatically (Excel-DNA, etc.)

### Build Steps

1. Open the `.csproj` file in the `Source/` folder of any add-in.
2. Restore NuGet packages.
3. Build in **Release** mode.
4. The output `.xll` files will be in `bin/Release/`.

---

## 📁 Repository Structure

```
├── README.md                              ← You are here
├── LICENSE
├── .gitignore
│
├── XLOOKUP-Excel-2019-2016-2013-2010/
│   ├── Source/    (C# source code)
│   └── Addin/    (XLookup_32bit.xll, XLookup_64bit.xll)
│
├── Excel-Crosshair-Highlighter/
│   ├── Source/    (C# source code)
│   └── Addin/    (ExcelCrosshairAddIn_32bit.xll, _64bit.xll)
│
├── Strikethrough-Filter-Excel-AddIn/
│   ├── Source/    (C# source code)
│   └── Addin/    (StrikethroughFilterAddIn_32bit.xll, _64bit.xll)
│
├── Excel-Delimiter-Text-Functions/
│   ├── Source/    (C# source code)
│   └── Addin/    (TextDelimiterAddIn_32bit.xll, _64bit.xll)
│
├── Excel-Filter-Copy/
│   ├── Source/    (C# source code)
│   └── Addin/    (ExcelFilterCopy_32bit.xll, _64bit.xll)
│
├── Excel-Merged-Filter-AddIn/
│   ├── Source/    (C# source code)
│   └── Addin/    (ExcelMergedFilter_32bit.xll, _64bit.xll)
│
├── Excel-Navigator-Arrows/
│   ├── Source/    (C# source code)
│   └── Addin/    (NavigatorArrowsAddIn_32bit.xll, _64bit.xll)
│
└── Replace-Many-Excel/
    ├── Source/    (C# source code)
    └── Addin/    (ReplaceManyAddIn_32bit.xll, _64bit.xll)
```

---

## 🔒 Troubleshooting

### "This add-in is not a valid Office Add-in"
You are loading the wrong bitness. Check your Excel version (32-bit vs 64-bit) and use the matching `.xll` file.

### "Security Warning: Application Add-ins have been disabled"
1. Right-click the downloaded `.xll` file.
2. Select **Properties**.
3. At the bottom, check **Unblock** → click **OK**.
4. Try loading in Excel again.

### Add-in does not appear after restart
1. Go to **File → Options → Add-ins → Manage: Excel Add-ins → Go...**
2. Ensure the add-in is checked.
3. If it says "not found", click **Browse...** and re-select the `.xll` file.

---

## 📄 License

[MIT License](LICENSE) — Free to use, modify, and distribute.

**Author:** Akshay Solanki
