# ExcelTab - In-Browser XLSX Spreadsheet Viewer (Chrome Extension)

ExcelTab intercepts spreadsheet downloads (`.xlsx`, `.xls`, `.csv`, `.ods`, `.tsv`) from web links and opens them instantly in a full-featured browser viewer tab without writing files to your Downloads folder.

---

## 🚀 Quick Installation Guide (Recommended)

### Method 1: Load Unpacked in Google Chrome / Edge / Brave

1. Download and extract **`exceltab-chrome-extension.zip`** to any folder on your computer.
2. Open Google Chrome and type `chrome://extensions/` in your address bar (or `edge://extensions/` in Microsoft Edge).
3. Toggle on **"Developer mode"** in the top-right corner.
4. Click the **"Load unpacked"** button in the top-left toolbar.
5. Select the extracted folder containing `manifest.json`.
6. That's it! ExcelTab is now active. Any spreadsheet link you click on any website will automatically open in a high-speed viewer tab.

---

### Method 2: Direct .CRX File Installation

1. Download **`exceltab-viewer.crx`**.
2. Open `chrome://extensions/` and ensure **Developer mode** is switched on.
3. Drag and drop the `exceltab-viewer.crx` file directly onto the `chrome://extensions/` page.
4. Click **"Add extension"** when prompted by Chrome.

---

## ⚡ Features

- **Zero-Download Workflow**: Intercepts links before your browser triggers a download.
- **Offline & Private**: All parsing happens 100% locally in your browser memory via WebAssembly/JS. No data is sent to external servers.
- **Full Spreadsheet Power**:
  - Multi-sheet tabs with row & column counts
  - Interactive formula bar with cell coordinates
  - Instant live search with highlighted matches
  - Quick column statistical calculations (Sum, Average, Min, Max, Count)
  - Zoom scaling (70% to 150%)
  - Clean Print & PDF export layout
  - One-click CSV export and original XLSX download option
  - **About Page & Modal**: Dedicated About page and interactive dialog with extension details and developer credits
- **Keyboard Bypass**: Hold `Shift` or `Alt` while clicking any link to trigger standard file download when desired.

---

## 👨‍💻 Developed By

**Beacon IT Team**  
Created with care for high-speed, secure, and privacy-focused in-browser document previewing.

