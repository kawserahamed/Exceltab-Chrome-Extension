/**
 * Standalone Viewer Controller for ExcelTab Pro
 * Preserves XLSX styles, colors, fonts, borders, widths, heights, merged cells, number formatting, freeze panes & PDF docs.
 */
(function() {
  const urlParams = new URLSearchParams(window.location.search);
  const targetFileId = urlParams.get('fileId');
  const targetUrl = urlParams.get('url');
  const targetName = urlParams.get('name') || 'Spreadsheet.xlsx';

  const fileNameEl = document.getElementById('fileName');
  const fileBadgeEl = document.getElementById('fileBadge');
  const sheetMetaEl = document.getElementById('sheetMeta');
  const loadingOverlay = document.getElementById('loadingOverlay');
  const loadingStatusText = document.getElementById('loadingStatusText');
  const gridContainer = document.getElementById('gridContainer');
  const gridHeader = document.getElementById('gridHeader');
  const gridBody = document.getElementById('gridBody');
  const sheetTabsContainer = document.getElementById('sheetTabsContainer');
  const cellAddressEl = document.getElementById('cellAddress');
  const cellFormulaEl = document.getElementById('cellFormula');
  const statsSummaryEl = document.getElementById('statsSummary');
  const searchInput = document.getElementById('searchInput');
  const searchCount = document.getElementById('searchCount');
  const btnZoomIn = document.getElementById('btnZoomIn');
  const btnZoomOut = document.getElementById('btnZoomOut');
  const zoomLevelEl = document.getElementById('zoomLevel');
  const btnExportCsv = document.getElementById('btnExportCsv');
  const btnPrint = document.getElementById('btnPrint');
  const btnDownloadOriginal = document.getElementById('btnDownloadOriginal');
  const btnAbout = document.getElementById('btnAbout');
  const aboutModal = document.getElementById('aboutModal');
  const btnCloseAboutModal = document.getElementById('btnCloseAboutModal');
  const btnCloseAboutModalBtn = document.getElementById('btnCloseAboutModalBtn');
  const btnOpenAboutPage = document.getElementById('btnOpenAboutPage');

  let currentWorkbook = null;
  let activeSheetName = '';
  let activeSheetWs = null;
  let activeSheetData = [];
  let currentZoom = 100;
  let selectedCell = { row: 0, col: 0 };
  let originalRawBuffer = null;
  let workbookFormat = null;      // XlsxFormat model of the open workbook (OOXML files only)
  let activeSheetFormat = null;   // formatting model of the active sheet, or null

  fileNameEl.textContent = targetName;
  document.title = targetName + ' - ExcelTab Pro Viewer';
  updateFileBadge(targetName);

  function updateFileBadge(filename) {
    if (!fileBadgeEl) return;
    const ext = filename.substring(filename.lastIndexOf('.') + 1).toUpperCase();
    fileBadgeEl.textContent = ext ? '.' + ext : '.DOCUMENT';
  }

  const EXCEL_THEME_PALETTE = [
    '#ffffff', '#000000', '#e7e6e6', '#44546a', '#5b9bd5',
    '#ed7d31', '#a5a5a5', '#ffc000', '#4472c4', '#70ad47'
  ];

  function parseColor(colorObj, defaultColor = null) {
    if (!colorObj) return defaultColor;
    if (colorObj.rgb) {
      let rgb = String(colorObj.rgb);
      if (rgb.length === 8) rgb = rgb.substring(2);
      if (rgb.length === 6) return '#' + rgb;
    }
    if (colorObj.theme !== undefined && EXCEL_THEME_PALETTE[colorObj.theme]) {
      let baseHex = EXCEL_THEME_PALETTE[colorObj.theme];
      if (colorObj.tint) {
        baseHex = applyTint(baseHex, colorObj.tint);
      }
      return baseHex;
    }
    return defaultColor;
  }

  function applyTint(hex, tint) {
    try {
      let num = parseInt(hex.replace('#', ''), 16);
      let r = (num >> 16), g = ((num >> 8) & 0x00FF), b = (num & 0x0000FF);
      if (tint > 0) {
        r = Math.round(r * (1 - tint) + 255 * tint);
        g = Math.round(g * (1 - tint) + 255 * tint);
        b = Math.round(b * (1 - tint) + 255 * tint);
      } else {
        r = Math.round(r * (1 + tint));
        g = Math.round(g * (1 + tint));
        b = Math.round(b * (1 + tint));
      }
      return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
    } catch (e) {
      return hex;
    }
  }

  function parseBorderSide(sideObj) {
    if (!sideObj || !sideObj.style) return null;
    const color = parseColor(sideObj.color, '#cbd5e1');
    let width = '1px';
    let style = 'solid';

    switch (sideObj.style) {
      case 'medium': width = '2px'; break;
      case 'thick': width = '3px'; break;
      case 'double': width = '3px'; style = 'double'; break;
      case 'dashed': case 'mediumDashed': style = 'dashed'; break;
      case 'dotted': case 'hair': style = 'dotted'; break;
      case 'dashDot': case 'slantDashDot': style = 'dashed'; break;
    }

    return `${width} ${style} ${color}`;
  }

  // Maps every merged cell to either its master (rendered with row/colspan) or a hidden slot.
  // Spans only count visible rows/columns, and the master moves to the first visible cell of the
  // merge so a merge whose top row or left column is hidden still renders like it does in Excel.
  function buildMergeMap(merges, rowHidden = [], colHidden = [], rowCount = Infinity, colCount = Infinity) {
    const map = {};
    if (!merges || !Array.isArray(merges)) return map;

    merges.forEach(m => {
      const endR = Math.min(m.e.r, rowCount - 1);
      const endC = Math.min(m.e.c, colCount - 1);
      const visRows = [];
      const visCols = [];
      for (let r = m.s.r; r <= endR; r++) if (!rowHidden[r]) visRows.push(r);
      for (let c = m.s.c; c <= endC; c++) if (!colHidden[c]) visCols.push(c);

      for (let r = m.s.r; r <= endR; r++) {
        for (let c = m.s.c; c <= endC; c++) {
          map[`${r},${c}`] = { hidden: true };
        }
      }

      if (visRows.length && visCols.length) {
        map[`${visRows[0]},${visCols[0]}`] = {
          master: true,
          rowspan: visRows.length,
          colspan: visCols.length,
          srcR: m.s.r,
          srcC: m.s.c,
          endR,
          endC,
          rows: visRows
        };
      }
    });

    return map;
  }

  function applyPrintMargins(margins) {
    let styleEl = document.getElementById('printMarginsStyle');
    if (!styleEl) {
      styleEl = document.createElement('style');
      styleEl.id = 'printMarginsStyle';
      document.head.appendChild(styleEl);
    }

    if (margins) {
      const top = margins.top !== undefined ? margins.top : 0.75;
      const bottom = margins.bottom !== undefined ? margins.bottom : 0.75;
      const left = margins.left !== undefined ? margins.left : 0.7;
      const right = margins.right !== undefined ? margins.right : 0.7;

      styleEl.textContent = `
        @media print {
          @page {
            margin: ${top}in ${right}in ${bottom}in ${left}in !important;
          }
        }
      `;
    } else {
      styleEl.textContent = '';
    }
  }

  // Paper sizes in inches (portrait) keyed by the OOXML pageSetup paperSize code.
  const PAPER_SIZES = {
    1: [8.5, 11], 3: [11, 17], 4: [17, 11], 5: [8.5, 14], 7: [7.25, 10.5], 8: [11.69, 16.54],
    9: [8.27, 11.69], 11: [5.83, 8.27], 12: [10.12, 14.33], 13: [7.17, 10.12]
  };

  // Mirrors the sheet's page layout (paper, orientation, scale, fit-to-width, print gridlines,
  // headings and centering) into print CSS so printing reproduces the workbook's page setup.
  function applyPrintLayout(sheetFormat, tableWidthPx) {
    let styleEl = document.getElementById('printLayoutStyle');
    if (!styleEl) {
      styleEl = document.createElement('style');
      styleEl.id = 'printLayoutStyle';
      document.head.appendChild(styleEl);
    }
    if (!sheetFormat) {
      styleEl.textContent = '';
      return;
    }

    const setup = sheetFormat.pageSetup || {};
    const options = sheetFormat.printOptions || {};
    const margins = sheetFormat.margins || { left: 0.7, right: 0.7 };
    const landscape = setup.orientation === 'landscape';
    const paper = PAPER_SIZES[setup.paperSize] || (setup.paperSize !== undefined ? PAPER_SIZES[1] : null);

    const pageRules = [];
    if (paper) {
      const [w, h] = landscape ? [paper[1], paper[0]] : paper;
      pageRules.push(`@page { size: ${w}in ${h}in; }`);
    } else if (landscape) {
      pageRules.push('@page { size: landscape; }');
    }

    let zoom = setup.scale ? setup.scale / 100 : 1;
    if (setup.fitToPage && setup.fitToWidth > 0 && tableWidthPx) {
      const pageWidthIn = paper ? (landscape ? paper[1] : paper[0]) : (landscape ? 11 : 8.5);
      const printableWidthPx = (pageWidthIn - margins.left - margins.right) * 96;
      zoom = Math.min(1, printableWidthPx / tableWidthPx);
    }

    const printRules = [
      `#spreadsheetTable.xf-mode { zoom: ${zoom.toFixed(3)}; transform: none !important; }`,
      '#spreadsheetTable.xf-mode th, #spreadsheetTable.xf-mode td { position: static !important; box-shadow: none !important; }'
    ];
    if (!options.gridLines) {
      printRules.push('#spreadsheetTable.xf-mode { --xg: transparent; }');
    }
    if (!options.headings) {
      printRules.push(
        '#spreadsheetTable.xf-mode thead th { height: 0 !important; padding: 0 !important; border: 0 !important; font-size: 0 !important; }',
        '#spreadsheetTable.xf-mode th.row-header, #spreadsheetTable.xf-mode th.corner-header { width: 0 !important; min-width: 0 !important; padding: 0 !important; border: 0 !important; font-size: 0 !important; }'
      );
      if (tableWidthPx) {
        printRules.push(`#spreadsheetTable.xf-mode { width: ${tableWidthPx - ROW_HEADER_PX}px !important; }`);
      }
    }
    if (options.horizontalCentered) {
      printRules.push('#spreadsheetTable.xf-mode { margin-left: auto; margin-right: auto; }');
    }

    styleEl.textContent = `${pageRules.join('\n')}\n@media print {\n  ${printRules.join('\n  ')}\n}`;
  }

  function createErrorDiagnostic(err, phase, extraInfo = {}) {
    const errorObj = {
      phase: phase || 'General Error',
      message: (err && err.message) ? err.message : String(err),
      name: (err && err.name) ? err.name : 'Error',
      stack: (err && err.stack) ? err.stack : 'No stack trace recorded',
      url: targetUrl || 'N/A',
      fileId: targetFileId || 'N/A',
      fileName: targetName || 'Spreadsheet.xlsx',
      timestamp: new Date().toISOString(),
      userAgent: navigator.userAgent,
      extra: extraInfo
    };

    try {
      chrome.storage.local.set({ last_exceltab_error: errorObj });
    } catch (e) {
      console.warn('Could not persist error diagnostic:', e);
    }

    return errorObj;
  }

  function showErrorUI(errorObj) {
    loadingOverlay.style.display = 'flex';

    const diagnosticText = [
      `=== ExcelTab Diagnostic Error Log ===`,
      `Timestamp: ${errorObj.timestamp}`,
      `Phase: ${errorObj.phase}`,
      `File Name: ${errorObj.fileName}`,
      `Target URL: ${errorObj.url}`,
      `File ID: ${errorObj.fileId}`,
      `Error Name: ${errorObj.name}`,
      `Error Message: ${errorObj.message}`,
      `User Agent: ${errorObj.userAgent}`,
      ``,
      `=== Stack Trace ===`,
      errorObj.stack
    ].join('\n');

    loadingOverlay.innerHTML = `
      <div style="max-width: 600px; width: 92%; background: #ffffff; border: 1px solid #fca5a5; border-radius: 8px; padding: 24px; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1); text-align: left; margin: auto;">
        <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 12px; border-bottom: 1px solid #e2e8f0; padding-bottom: 12px;">
          <span style="font-size: 26px;">⚠️</span>
          <div>
            <h3 style="font-size: 16px; font-weight: 700; color: #dc2626; margin: 0;">ExcelTab Error Diagnostic</h3>
            <span style="font-size: 11px; color: #64748b;">Phase: <strong style="color:#0f172a;">${escapeHtml(errorObj.phase)}</strong></span>
          </div>
        </div>

        <div style="background: #fef2f2; border-left: 4px solid #ef4444; padding: 12px; border-radius: 4px; margin-bottom: 16px;">
          <p style="font-size: 12px; font-weight: 600; color: #991b1b; margin-bottom: 4px;">Error Message:</p>
          <p style="font-size: 12px; color: #b91c1c; font-family: monospace; word-break: break-word; margin: 0;">${escapeHtml(errorObj.message)}</p>
        </div>

        <details style="background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 4px; padding: 10px; margin-bottom: 16px; font-size: 11px;">
          <summary style="cursor: pointer; font-weight: 600; color: #059669; outline: none;">🔍 View Technical Diagnostics & Full Stack Trace</summary>
          <pre style="margin-top: 10px; color: #334155; font-family: monospace; white-space: pre-wrap; word-break: break-all; max-height: 180px; overflow-y: auto; font-size: 11px; line-height: 1.4;">${escapeHtml(diagnosticText)}</pre>
        </details>

        <div style="display: flex; flex-wrap: wrap; gap: 10px; justify-content: flex-end; border-top: 1px solid #e2e8f0; padding-top: 14px;">
          <button id="btnCopyLog" class="btn btn-secondary" style="font-size: 12px;">📋 Copy Error Log</button>
          ${targetUrl ? `<button id="btnDirectDownloadFallback" class="btn btn-primary" style="font-size: 12px;">⬇️ Download File Directly</button>` : ''}
          <button id="retrySample" class="btn btn-secondary" style="font-size: 12px;">🧪 Load Sample Sheet</button>
        </div>
      </div>
    `;

    const copyBtn = document.getElementById('btnCopyLog');
    if (copyBtn) {
      copyBtn.onclick = () => {
        navigator.clipboard.writeText(diagnosticText).then(() => {
          copyBtn.textContent = '✅ Copied!';
          setTimeout(() => copyBtn.textContent = '📋 Copy Error Log', 2000);
        }).catch(e => {
          console.error('Copy failed:', e);
        });
      };
    }

    const fallbackBtn = document.getElementById('btnDirectDownloadFallback');
    if (fallbackBtn && targetUrl) {
      fallbackBtn.onclick = () => {
        chrome.runtime.sendMessage({
          action: 'BYPASS_AND_DOWNLOAD',
          url: targetUrl,
          filename: targetName
        });
      };
    }

    const retryBtn = document.getElementById('retrySample');
    if (retryBtn) {
      retryBtn.onclick = () => loadSampleSpreadsheet();
    }
  }

  // Load workbook or PDF document from local storage or URL
  async function loadSpreadsheet() {
    if (targetFileId) {
      loadingStatusText.textContent = 'Loading document data from storage...';
      chrome.storage.local.get(targetFileId, (items) => {
        const entry = items[targetFileId];
        if (entry && entry.data) {
          try {
            const binaryString = atob(entry.data);
            const len = binaryString.length;
            const bytes = new Uint8Array(len);
            for (let i = 0; i < len; i++) {
              bytes[i] = binaryString.charCodeAt(i);
            }
            originalRawBuffer = bytes.buffer;
            renderWorkbook(bytes.buffer);
            chrome.storage.local.remove(targetFileId);
          } catch (e) {
            console.error('Failed to parse stored binary data:', e);
            const diag = createErrorDiagnostic(e, 'Local Storage Base64 Decode');
            showErrorUI(diag);
          }
        } else {
          const diag = createErrorDiagnostic(
            new Error('Document data expired or was not found in local storage for ID: ' + targetFileId),
            'Local Storage Retrieval'
          );
          showErrorUI(diag);
        }
      });
      return;
    }

    if (!targetUrl) {
      const diag = createErrorDiagnostic(
        new Error('No target URL or file ID parameter was provided to viewer.'),
        'Initialization'
      );
      showErrorUI(diag);
      return;
    }

    let arrayBuffer;
    let directFetchErr = null;

    try {
      loadingStatusText.textContent = 'Fetching file directly...';
      const response = await fetch(targetUrl);
      if (!response.ok) {
        throw new Error(`HTTP Error ${response.status}: ${response.statusText}`);
      }
      arrayBuffer = await response.arrayBuffer();
    } catch (fetchErr) {
      directFetchErr = fetchErr;
      console.warn('Direct fetch failed, attempting service worker proxy...', fetchErr);
    }

    if (!arrayBuffer) {
      try {
        loadingStatusText.textContent = 'Fetching via background service worker proxy...';
        arrayBuffer = await fetchViaBackground(targetUrl);
      } catch (bgErr) {
        console.error('Background fetch proxy failed:', bgErr);
        const combinedMsg = `Direct fetch failed: ${directFetchErr ? directFetchErr.message : 'Unknown error'}. Background proxy failed: ${bgErr.message}`;
        const diag = createErrorDiagnostic(new Error(combinedMsg), 'Network Fetch', {
          directError: directFetchErr ? directFetchErr.stack : null,
          bgError: bgErr.stack
        });
        showErrorUI(diag);
        return;
      }
    }

    originalRawBuffer = arrayBuffer;
    renderWorkbook(arrayBuffer);
  }

  function fetchViaBackground(url) {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendMessage({ action: 'FETCH_SPREADSHEET', url }, (response) => {
        if (chrome.runtime.lastError) {
          return reject(new Error('Chrome Messaging Error: ' + chrome.runtime.lastError.message));
        }
        if (!response || !response.success) {
          return reject(new Error((response && response.error) || 'Background proxy returned no data.'));
        }
        try {
          const binaryString = atob(response.data);
          const len = binaryString.length;
          const bytes = new Uint8Array(len);
          for (let i = 0; i < len; i++) {
            bytes[i] = binaryString.charCodeAt(i);
          }
          resolve(bytes.buffer);
        } catch (decodeErr) {
          reject(new Error('Failed to decode proxy Base64 response: ' + decodeErr.message));
        }
      });
    });
  }

  function loadSampleSpreadsheet() {
    fileNameEl.textContent = 'Sample_Financial_Report.xlsx';
    document.title = 'Sample_Financial_Report.xlsx - ExcelTab Pro Viewer';
    updateFileBadge('Sample_Financial_Report.xlsx');
    const sampleData = [
      ['Region', 'Q1 Sales ($)', 'Q2 Sales ($)', 'Q3 Sales ($)', 'Q4 Sales ($)', 'Total ($)'],
      ['North America', 15000, 18000, 21000, 25000, 79000],
      ['Europe', 12000, 14500, 16000, 19000, 61500],
      ['Asia Pacific', 22000, 26000, 31000, 35000, 114000],
      ['Latin America', 8000, 9500, 11000, 13000, 41500],
      ['Total', 57000, 68000, 79000, 92000, 296000]
    ];
    const ws = XLSX.utils.aoa_to_sheet(sampleData);

    // Apply sample header styling
    for (let c = 0; c < 6; c++) {
      const ref = XLSX.utils.encode_cell({ r: 0, c });
      if (ws[ref]) {
        ws[ref].s = {
          fill: { fgColor: { rgb: '10B981' } },
          font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
          alignment: { horizontal: c > 0 ? 'right' : 'left' }
        };
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sales Summary');

    currentWorkbook = wb;
    workbookFormat = null;
    buildSheetTabs(wb.SheetNames);
    selectSheet(wb.SheetNames[0]);
    loadingOverlay.style.display = 'none';
  }

  function isPdfBuffer(buffer) {
    if (targetName && targetName.toLowerCase().endsWith('.pdf')) return true;
    if (!buffer) return false;
    const bytes = new Uint8Array(buffer);
    if (bytes.length < 5) return false;
    return bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46 && bytes[4] === 0x2D;
  }

  function renderPdfDocument(data) {
    updateFileBadge(targetName || 'document.pdf');
    if (sheetMetaEl) sheetMetaEl.textContent = '(PDF Document)';

    const formulaBar = document.querySelector('.formula-bar');
    if (formulaBar) formulaBar.style.display = 'none';

    const sheetTabsBar = document.querySelector('.sheet-tabs-bar');
    if (sheetTabsBar) sheetTabsBar.style.display = 'none';

    let pdfSrc = targetUrl;
    if (data && data.byteLength > 0) {
      const blob = new Blob([data], { type: 'application/pdf' });
      pdfSrc = URL.createObjectURL(blob);
    }

    gridContainer.innerHTML = `
      <embed src="${escapeHtml(pdfSrc)}" type="application/pdf" style="width: 100%; height: 100%; border: none;" />
    `;

    loadingOverlay.style.display = 'none';
  }

  function isHtmlBuffer(buffer) {
    const bytes = new Uint8Array(buffer);
    if (bytes.length < 4) return false;

    if (bytes[0] === 0x50 && bytes[1] === 0x4B && bytes[2] === 0x03 && bytes[3] === 0x04) {
      return false;
    }
    if (bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0) {
      return false;
    }

    let sample = '';
    const sampleLen = Math.min(bytes.length, 512);
    for (let i = 0; i < sampleLen; i++) {
      sample += String.fromCharCode(bytes[i]);
    }
    sample = sample.trim().toLowerCase();
    return sample.startsWith('<!doctype html') || sample.startsWith('<html') || sample.startsWith('<head') || sample.includes('<body');
  }

  function renderWorkbook(data) {
    loadingStatusText.textContent = 'Parsing file data & formatting...';
    try {
      if (!data || data.byteLength === 0) {
        throw new Error('Downloaded file payload is 0 bytes (Empty File).');
      }

      if (isPdfBuffer(data)) {
        renderPdfDocument(data);
        return;
      }

      if (typeof XLSX === 'undefined') {
        throw new Error('XLSX library failed to initialize or window.XLSX is undefined.');
      }

      if (isHtmlBuffer(data)) {
        throw new Error('The web server returned an HTML webpage or login screen instead of a binary file.');
      }

      const wb = XLSX.read(data, {
        type: 'array',
        cellStyles: true,
        cellFormulas: true,
        cellDates: true,
        cellNF: true,
        sheetStubs: true
      });

      currentWorkbook = wb;
      if (!wb.SheetNames || wb.SheetNames.length === 0) {
        throw new Error('Workbook parsed successfully but contains no worksheets.');
      }

      // Read the full original formatting (styles, theme colors, hidden rows/cols, layout).
      // Any failure here falls back to the basic SheetJS rendering instead of breaking the view.
      workbookFormat = null;
      if (typeof XlsxFormat !== 'undefined') {
        try {
          workbookFormat = XlsxFormat.parse(data);
        } catch (fmtErr) {
          console.warn('Could not read XLSX formatting, using basic rendering:', fmtErr);
        }
      }

      buildSheetTabs(wb.SheetNames);
      selectSheet(initialSheetName(wb));
      loadingOverlay.style.display = 'none';
    } catch (e) {
      console.error('Workbook render error:', e);
      const diag = createErrorDiagnostic(e, 'File Parsing');
      showErrorUI(diag);
    }
  }

  function isSheetHidden(name) {
    if (workbookFormat) return !!workbookFormat.hiddenSheets[name];
    // Other formats (XLS, ODS...): SheetJS reports visibility as Hidden 1 (hidden) / 2 (very hidden)
    const sheetsMeta = currentWorkbook && currentWorkbook.Workbook && currentWorkbook.Workbook.Sheets;
    const index = currentWorkbook ? currentWorkbook.SheetNames.indexOf(name) : -1;
    return !!(sheetsMeta && sheetsMeta[index] && sheetsMeta[index].Hidden);
  }

  // Sheets hidden in the workbook stay hidden, unless every sheet is hidden.
  function visibleSheetNames(sheetNames) {
    const visible = sheetNames.filter(name => !isSheetHidden(name));
    return visible.length ? visible : sheetNames;
  }

  // Open on the sheet that was active when the workbook was saved.
  function initialSheetName(wb) {
    const visible = visibleSheetNames(wb.SheetNames);
    if (workbookFormat) {
      const saved = workbookFormat.sheetOrder[workbookFormat.activeTab];
      if (saved && visible.indexOf(saved) >= 0) return saved;
    }
    return visible[0];
  }

  function buildSheetTabs(sheetNames) {
    sheetTabsContainer.innerHTML = '';
    visibleSheetNames(sheetNames).forEach(name => {
      const btn = document.createElement('button');
      btn.className = 'sheet-tab' + (name === activeSheetName ? ' active' : '');
      btn.textContent = name;
      btn.onclick = () => selectSheet(name);
      const sheetFormat = workbookFormat && workbookFormat.sheets[name];
      if (sheetFormat && sheetFormat.tabColor) {
        btn.classList.add('has-tab-color');
        btn.style.setProperty('--tab-color', sheetFormat.tabColor);
      }
      sheetTabsContainer.appendChild(btn);
    });
  }

  function selectSheet(name) {
    activeSheetName = name;
    const ws = currentWorkbook.Sheets[name];
    if (!ws) return;
    activeSheetWs = ws;
    activeSheetFormat = (workbookFormat && workbookFormat.sheets[name]) || null;

    // Update active tab styles
    Array.from(sheetTabsContainer.children).forEach(tab => {
      tab.classList.toggle('active', tab.textContent === name);
    });

    activeSheetData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    renderGrid(ws);
  }

  function colIndexToName(index) {
    let name = '';
    let num = index;
    while (num >= 0) {
      name = String.fromCharCode((num % 26) + 65) + name;
      num = Math.floor(num / 26) - 1;
    }
    return name;
  }

  const ROW_HEADER_PX = 46;          // row number column width (matches viewer.css)
  const LEGACY_DEFAULT_COL_PX = 100;
  const LEGACY_DEFAULT_ROW_PX = 26;
  const SPILL_MAX_COLS = 30;         // how far overflowing text may run into empty neighbours
  const FREEZE_LINE_COLOR = '#9ca3af';

  // Sizing/visibility source for files without an OOXML formatting model (CSV, XLS, sample sheet).
  function legacyLayout(ws) {
    const cols = ws['!cols'] || [];
    const rows = ws['!rows'] || [];
    let frozenRows = 0;
    let frozenCols = 0;
    if (ws['!freeze']) {
      frozenCols = ws['!freeze'].xSplit || 0;
      frozenRows = ws['!freeze'].ySplit || 0;
    } else if (ws['!views'] && ws['!views'][0]) {
      frozenCols = ws['!views'][0].xSplit || 0;
      frozenRows = ws['!views'][0].ySplit || 0;
    }
    return {
      fidelity: false,
      frozenRows,
      frozenCols,
      defaultColPx: LEGACY_DEFAULT_COL_PX,
      defaultRowPx: LEGACY_DEFAULT_ROW_PX,
      colPx(c) {
        const col = cols[c];
        if (!col) return null;
        if (col.wpx) return col.wpx;
        if (col.wch) return Math.round(col.wch * 8 + 12);
        if (col.width) return Math.round(col.width * 8 + 12);
        return null;
      },
      colHidden(c) {
        return !!(cols[c] && cols[c].hidden);
      },
      rowPx(r) {
        const row = rows[r];
        if (!row) return null;
        // SheetJS reports hpx equal to the point value, so convert from points when available.
        if (row.hpt) return Math.round(row.hpt * 96 / 72);
        if (row.hpx) return row.hpx;
        return null;
      },
      rowHidden(r) {
        return !!(rows[r] && rows[r].hidden);
      }
    };
  }

  // Sizing/visibility source backed by the original workbook's sheet XML.
  function fidelityLayout(sf) {
    return {
      fidelity: true,
      frozenRows: sf.frozenRows,
      frozenCols: sf.frozenCols,
      defaultColPx: sf.defaultColPx,
      defaultRowPx: sf.defaultRowPx,
      colPx(c) {
        const col = XlsxFormat.colInfo(sf, c);
        return col && col.px !== null ? col.px : sf.defaultColPx;
      },
      colHidden(c) {
        const col = XlsxFormat.colInfo(sf, c);
        return !!(col && col.hidden);
      },
      rowPx(r) {
        const row = sf.rows.get(r);
        return row && row.px !== null ? row.px : sf.defaultRowPx;
      },
      rowHidden(r) {
        const row = sf.rows.get(r);
        return row ? row.hidden : sf.zeroHeight;
      }
    };
  }

  function cellDisplayText(cell) {
    if (!cell) return '';
    // cell.w contains the formatted currency/date/percentage/number string
    if (cell.w !== undefined && cell.w !== null) return String(cell.w);
    if (cell.v !== undefined && cell.v !== null) return String(cell.v);
    return '';
  }

  function isBlankCell(cell) {
    return !cell || cell.t === 'z' || cell.v === undefined || cell.v === null || cell.v === '';
  }

  // Styles for cells rendered from SheetJS data alone (no OOXML formatting model).
  function legacyCellStyles(cell, inlineStyles) {
    let hasFill = false;
    if (cell.s) {
      const s = cell.s;

      // Fill / Background Color. The SheetJS community build stores the fill itself in cell.s
      // ({ patternType, fgColor, bgColor }); styles built in code use { fill: {...} }.
      const fill = s.fill || ((s.patternType || s.fgColor || s.bgColor) ? s : null);
      if (fill && fill.patternType !== 'none') {
        const bg = parseColor(fill.fgColor || fill.bgColor);
        if (bg) {
          inlineStyles.push(`background-color: ${bg};`);
          hasFill = true;
        }
      }

      // Font Properties
      if (s.font) {
        const f = s.font;
        if (f.bold) inlineStyles.push('font-weight: bold;');
        if (f.italic) inlineStyles.push('font-style: italic;');
        if (f.underline) inlineStyles.push('text-decoration: underline;');
        if (f.sz) inlineStyles.push(`font-size: ${f.sz}pt;`);
        if (f.name) inlineStyles.push(`font-family: '${String(f.name).replace(/["'\\;<>{}&]/g, '')}', sans-serif;`);
        const fontColor = parseColor(f.color);
        if (fontColor) inlineStyles.push(`color: ${fontColor};`);
      }

      // Alignment
      if (s.alignment) {
        const a = s.alignment;
        if (a.horizontal) inlineStyles.push(`text-align: ${a.horizontal};`);
        if (a.vertical) {
          const vMap = { top: 'top', center: 'middle', bottom: 'bottom' };
          inlineStyles.push(`vertical-align: ${vMap[a.vertical] || a.vertical};`);
        }
        if (a.wrapText) inlineStyles.push('white-space: normal; word-break: break-word;');
      }

      // Borders
      if (s.border) {
        const b = s.border;
        const topB = parseBorderSide(b.top);
        const botB = parseBorderSide(b.bottom);
        const leftB = parseBorderSide(b.left);
        const rightB = parseBorderSide(b.right);

        if (topB) inlineStyles.push(`border-top: ${topB};`);
        if (botB) inlineStyles.push(`border-bottom: ${botB};`);
        if (leftB) inlineStyles.push(`border-left: ${leftB};`);
        if (rightB) inlineStyles.push(`border-right: ${rightB};`);
      }
    }

    // Default numeric alignment if no explicit alignment given
    if ((!cell.s || !cell.s.alignment || !cell.s.alignment.horizontal) && (cell.t === 'n' || cell.t === 'd' || typeof cell.v === 'number')) {
      inlineStyles.push('text-align: right; font-variant-numeric: tabular-nums;');
    }
    return hasFill;
  }

  // Rich-text runs (mixed fonts/colors inside one cell). A run with properties defines its
  // font completely, so reset bold/italic/decoration before applying its own settings.
  function richTextHtml(runs) {
    return runs.map(run => {
      let css = '';
      if (run.font) {
        css = 'font-weight:normal;font-style:normal;text-decoration:none;' + XlsxFormat.fontCss(run.font, null);
        if (run.font.vertAlign === 'superscript') css += 'vertical-align:super;font-size:smaller;';
        else if (run.font.vertAlign === 'subscript') css += 'vertical-align:sub;font-size:smaller;';
      }
      const text = escapeHtml(run.text);
      return css ? `<span style="${css}">${text}</span>` : `<span>${text}</span>`;
    }).join('');
  }

  let measureCtx = null;
  function textWidthPx(text, font, baseFont) {
    if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
    const size = font.sz || baseFont.sz || 11;
    const family = XlsxFormat.fontStack(font.name || baseFont.name) || 'sans-serif';
    measureCtx.font = `${font.i ? 'italic ' : ''}${font.b ? 'bold ' : ''}${size}pt ${family}`;
    return measureCtx.measureText(text).width;
  }

  // Excel shows "####" instead of a formatted number/date that does not fit its column.
  function fitNumberText(text, cell, xf, availablePx, baseFont) {
    if (!text || !cell || (cell.t !== 'n' && cell.t !== 'd') || !cell.z || String(cell.z) === 'General') return text;
    if (xf.align.rot || xf.align.wrap || xf.align.shrink) return text;
    if (textWidthPx(text, xf.font, baseFont) <= availablePx) return text;
    return '#'.repeat(Math.max(1, Math.floor(availablePx / textWidthPx('#', xf.font, baseFont))));
  }

  function rotationWrap(html, rot) {
    if (!rot) return html;
    if (rot === 255) {
      return `<span style="display:inline-block; writing-mode:vertical-rl; text-orientation:upright;">${html}</span>`;
    }
    const deg = rot <= 90 ? -rot : rot - 90;
    return `<span style="display:inline-block; transform:rotate(${deg}deg); transform-origin:center;">${html}</span>`;
  }

  function applyTableMode(table, sf) {
    table.classList.toggle('xf-mode', !!sf);
    table.classList.toggle('xf-nogrid', !!(sf && !sf.showGrid));
    if (sf) {
      table.style.setProperty('--xf-font', XlsxFormat.fontStack(sf.baseFont.name) || 'sans-serif');
      table.style.setProperty('--xf-size', (sf.baseFont.sz || 11) + 'pt');
      table.style.setProperty('--xf-color', sf.baseFont.color || '#000000');
    } else {
      table.style.removeProperty('--xf-font');
      table.style.removeProperty('--xf-size');
      table.style.removeProperty('--xf-color');
      table.style.width = '';
    }
  }

  function renderGrid(ws) {
    const table = document.getElementById('spreadsheetTable');
    const sf = activeSheetFormat;
    applyTableMode(table, sf);

    if (!ws || !ws['!ref']) {
      gridHeader.innerHTML = '';
      gridBody.innerHTML = '<tr><td style="padding: 32px; text-align: center; color: #94a3b8;">This sheet is empty.</td></tr>';
      table.style.width = '';
      applyPrintLayout(null);
      return;
    }

    const range = XLSX.utils.decode_range(ws['!ref']);
    const rowCount = range.e.r + 1;
    const colCount = range.e.c + 1;

    const sheetCount = currentWorkbook ? visibleSheetNames(currentWorkbook.SheetNames).length : 1;
    sheetMetaEl.textContent = `(${sheetCount} sheet${sheetCount > 1 ? 's' : ''} • ${rowCount} rows × ${colCount} cols)`;

    const layout = sf ? fidelityLayout(sf) : legacyLayout(ws);

    // Column widths & visibility (hidden columns are never rendered)
    const colPx = new Array(colCount);
    const colHidden = new Array(colCount);
    const visCols = [];
    for (let c = 0; c < colCount; c++) {
      colHidden[c] = layout.colHidden(c);
      colPx[c] = layout.colPx(c);
      if (!colHidden[c]) visCols.push(c);
    }

    // Row visibility (hidden rows are never rendered)
    const rowHidden = new Array(rowCount);
    const visRows = [];
    for (let r = 0; r < rowCount; r++) {
      rowHidden[r] = layout.rowHidden(r);
      if (!rowHidden[r]) visRows.push(r);
    }

    // Next visible column/row after a given index (-1 when none) for border sharing & overflow
    const nextVisCol = new Array(colCount).fill(-1);
    for (let i = visCols.length - 2; i >= 0; i--) {
      for (let c = visCols[i]; c < visCols[i + 1]; c++) nextVisCol[c] = visCols[i + 1];
    }
    const nextVisRow = new Array(rowCount).fill(-1);
    for (let i = visRows.length - 2; i >= 0; i--) {
      for (let r = visRows[i]; r < visRows[i + 1]; r++) nextVisRow[r] = visRows[i + 1];
    }

    // Build merged cells lookup map
    const mergeMap = buildMergeMap(ws['!merges'], rowHidden, colHidden, rowCount, colCount);

    const frozenRows = layout.frozenRows;
    const frozenCols = layout.frozenCols;
    const lastFrozenRow = visRows.filter(r => r < frozenRows).pop();
    const lastFrozenCol = visCols.filter(c => c < frozenCols).pop();

    const styleAt = (r, c) => XlsxFormat.styleAt(sf, r, c);
    const cellAt = (r, c) => ws[XLSX.utils.encode_cell({ r, c })];

    // Build Header Row
    let headerHtml = '<tr><th class="corner-header">#</th>';
    const colLeft = [];
    let tableWidth = ROW_HEADER_PX;
    for (const c of visCols) {
      const colWidth = colPx[c];
      colLeft[c] = tableWidth;
      let widthStyle = colWidth ? `width: ${colWidth}px; min-width: ${colWidth}px; max-width: ${colWidth}px;` : '';
      if (c < frozenCols) widthStyle += ` position: sticky; left: ${tableWidth}px; z-index: 25;`;
      headerHtml += `<th data-c="${c}" style="${widthStyle}">${colIndexToName(c)}</th>`;
      tableWidth += colWidth || layout.defaultColPx;
    }
    headerHtml += '</tr>';
    gridHeader.innerHTML = headerHtml;
    table.style.width = sf ? `${tableWidth}px` : '';

    // Build Body Rows
    let bodyHtml = '';
    let frozenTop = 0; // offset of the next frozen row below the column header

    for (const r of visRows) {
      const rowHeight = layout.rowPx(r);
      const rowHeightStyle = rowHeight ? `height: ${rowHeight}px;` : '';
      const isFrozenRow = r < frozenRows;
      const stickyTop = `top: calc(var(--xt-header-h, 27px) + ${frozenTop}px);`;

      bodyHtml += `<tr data-r="${r}" style="${rowHeightStyle}">`;
      bodyHtml += `<th class="row-header" style="${rowHeightStyle}${isFrozenRow ? ` ${stickyTop} z-index: 17;` : ''}">${r + 1}</th>`;

      for (const c of visCols) {
        const mergeInfo = mergeMap[`${r},${c}`];
        if (mergeInfo && mergeInfo.hidden) {
          // Skip hidden merged cell
          continue;
        }

        // A merged area shows the content and format of its top-left cell
        const srcR = mergeInfo ? mergeInfo.srcR : r;
        const srcC = mergeInfo ? mergeInfo.srcC : c;
        const endR = mergeInfo ? mergeInfo.endR : r;
        const endC = mergeInfo ? mergeInfo.endC : c;
        const cell = cellAt(srcR, srcC);
        const cellValue = cellDisplayText(cell);
        const inlineStyles = [];
        let cellClass = '';
        let content;
        let hasFill = false;

        if (sf) {
          const xf = styleAt(srcR, srcC);
          inlineStyles.push(xf.css);
          if (xf.fill) {
            inlineStyles.push(xf.fill.css);
            hasFill = true;
          }

          // "General" horizontal alignment depends on the value type, like Excel
          const h = xf.align.h;
          if ((!h || h === 'general') && cell) {
            if (cell.t === 'n' || cell.t === 'd') inlineStyles.push('text-align: right;');
            else if (cell.t === 'b' || cell.t === 'e') inlineStyles.push('text-align: center;');
          }

          // Number format colors, e.g. [Red] for negatives
          if (cell && cell.z && !isBlankCell(cell)) {
            const fmtColor = XlsxFormat.numberFormatColor(String(cell.z), cell.v, sf.indexed);
            if (fmtColor) inlineStyles.push(`color: ${fmtColor};`);
          }

          // Borders: a shared edge uses this cell's side or the neighbour's facing side.
          // Unbordered edges of filled cells take the fill color so gridlines don't show through.
          const nextC = nextVisCol[endC];
          const nextR = nextVisRow[endR];
          const rightXf = nextC >= 0 ? styleAt(r, nextC) : null;
          const belowXf = nextR >= 0 ? styleAt(nextR, c) : null;
          const ownRight = (endC === srcC ? xf : styleAt(srcR, endC)).border.right;
          const ownBottom = (endR === srcR ? xf : styleAt(endR, srcC)).border.bottom;
          const right = ownRight || (rightXf && rightXf.border.left);
          const bottom = ownBottom || (belowXf && belowXf.border.top);
          if (right) {
            inlineStyles.push(`border-right: ${right};`);
          } else {
            const edgeFill = (xf.fill && xf.fill.color) || (rightXf && rightXf.fill && rightXf.fill.color);
            if (edgeFill) inlineStyles.push(`border-right-color: ${edgeFill};`);
          }
          if (bottom) {
            inlineStyles.push(`border-bottom: ${bottom};`);
          } else {
            const edgeFill = (xf.fill && xf.fill.color) || (belowXf && belowXf.fill && belowXf.fill.color);
            if (edgeFill) inlineStyles.push(`border-bottom-color: ${edgeFill};`);
          }
          if (c === visCols[0] && xf.border.left) inlineStyles.push(`border-left: ${xf.border.left};`);
          if (r === visRows[0] && xf.border.top) inlineStyles.push(`border-top: ${xf.border.top};`);

          // Content box: clipped to the row height like Excel
          let boxHeight = rowHeight;
          if (mergeInfo) {
            boxHeight = mergeInfo.rows.reduce((sum, mr) => sum + (layout.rowPx(mr) || layout.defaultRowPx), 0);
          }
          const innerStyles = [`max-height: ${Math.max(boxHeight - 1, 0)}px;`];

          // Unwrapped text overflows into empty cells to its right, like Excel
          const canSpill = !mergeInfo && cellValue && cell && (cell.t === 's' || cell.t === 'str') &&
            !xf.align.wrap && !xf.align.rot && !xf.align.shrink && (!h || h === 'general' || h === 'left');
          // On screen the overflow stops at the freeze line (as in Excel); print ignores panes.
          if (canSpill) {
            let spillWidth = colPx[c];
            let printWidth = colPx[c];
            let k = nextVisCol[c];
            let steps = 0;
            let printSteps = 0;
            let crossedFreeze = false;
            while (k >= 0 && printSteps < SPILL_MAX_COLS && !mergeMap[`${r},${k}`] && isBlankCell(cellAt(r, k))) {
              if ((k < frozenCols) !== (c < frozenCols)) crossedFreeze = true;
              printWidth += colPx[k];
              printSteps++;
              if (!crossedFreeze) {
                spillWidth += colPx[k];
                steps++;
              }
              k = nextVisCol[k];
            }
            if (printSteps) {
              cellClass = steps ? ' class="xs"' : ' class="xsp"';
              if (steps) innerStyles.push(`width: ${spillWidth - 7}px;`);
              if (printSteps !== steps) innerStyles.push(`--xw-print: ${printWidth - 7}px;`);
            }
          }

          const runs = sf.rich.get(`${srcR},${srcC}`);
          let displayValue = cellValue;
          if (!mergeInfo) {
            const indentPx = xf.align.indent ? Math.round(xf.align.indent * 9) : 0;
            displayValue = fitNumberText(cellValue, cell, xf, colPx[c] - 7 - indentPx, sf.baseFont);
          }
          const html = runs ? richTextHtml(runs) : escapeHtml(displayValue);
          content = `<div class="xc" style="${innerStyles.join(' ')}">${rotationWrap(html, xf.align.rot)}</div>`;
        } else {
          const colWidth = colPx[c];
          if (colWidth) {
            inlineStyles.push(`width: ${colWidth}px; min-width: ${colWidth}px; max-width: ${colWidth}px;`);
          }
          if (cell) hasFill = legacyCellStyles(cell, inlineStyles);
          content = escapeHtml(cellValue);
        }

        // Freeze panes: frozen rows stick below the header, frozen columns beside the row numbers
        const isFrozenCol = c < frozenCols;
        if (isFrozenRow || isFrozenCol) {
          inlineStyles.push('position: sticky;');
          if (isFrozenRow) inlineStyles.push(stickyTop);
          if (isFrozenCol) inlineStyles.push(`left: ${colLeft[c]}px;`);
          inlineStyles.push(`z-index: ${isFrozenRow && isFrozenCol ? 16 : isFrozenRow ? 14 : 12};`);
          if (!hasFill) inlineStyles.push('background-color: #ffffff;');
          const shadows = [];
          if (r === lastFrozenRow || (mergeInfo && mergeInfo.rows.indexOf(lastFrozenRow) >= 0)) shadows.push(`inset 0 -1px 0 ${FREEZE_LINE_COLOR}`);
          if (c === lastFrozenCol) shadows.push(`inset -1px 0 0 ${FREEZE_LINE_COLOR}`);
          if (shadows.length) inlineStyles.push(`box-shadow: ${shadows.join(', ')};`);
        }

        let mergeAttrs = '';
        if (mergeInfo && mergeInfo.master) {
          if (mergeInfo.rowspan > 1) mergeAttrs += ` rowspan="${mergeInfo.rowspan}"`;
          if (mergeInfo.colspan > 1) mergeAttrs += ` colspan="${mergeInfo.colspan}"`;
        }

        bodyHtml += `<td data-r="${srcR}" data-c="${srcC}"${mergeAttrs}${cellClass} style="${inlineStyles.join(' ')}">${content}</td>`;
      }

      bodyHtml += '</tr>';
      if (isFrozenRow) frozenTop += rowHeight || layout.defaultRowPx;
    }

    gridBody.innerHTML = bodyHtml;
    table.style.setProperty('--xt-header-h', (gridHeader.offsetHeight || 27) + 'px');

    // Apply Page Margins & page setup for Print
    applyPrintMargins((sf && sf.margins) || ws['!margins']);
    applyPrintLayout(sf, sf ? tableWidth : 0);

    selectCell(visRows.length ? visRows[0] : 0, visCols.length ? visCols[0] : 0);
  }

  function escapeHtml(str) {
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function selectCell(r, c) {
    selectedCell = { row: r, col: c };

    // 1. Remove previous selection & active header highlights
    const prevCell = gridBody.querySelector('.selected');
    if (prevCell) prevCell.classList.remove('selected');

    const prevActiveHeaders = gridContainer.querySelectorAll('.header-active');
    prevActiveHeaders.forEach(el => el.classList.remove('header-active'));

    // 2. Add selection to target cell
    const targetCell = gridBody.querySelector('td[data-r="' + r + '"][data-c="' + c + '"]');
    if (targetCell) {
      targetCell.classList.add('selected');
    }

    // 3. Highlight corresponding column & row headers (Excel / Google Sheets style)
    // (looked up by index attributes because hidden rows/columns are not rendered)
    const headerCol = gridHeader.querySelector('th[data-c="' + c + '"]');
    if (headerCol) {
      headerCol.classList.add('header-active');
    }

    const targetRow = gridBody.querySelector('tr[data-r="' + r + '"]');
    if (targetRow) {
      const rowHeader = targetRow.querySelector('th.row-header');
      if (rowHeader) rowHeader.classList.add('header-active');
    }

    // 4. Update address & formula bar
    const addr = colIndexToName(c) + (r + 1);
    cellAddressEl.textContent = addr;

    const cellRef = XLSX.utils.encode_cell({ r, c });
    const cell = activeSheetWs ? activeSheetWs[cellRef] : null;

    let cellFormula = '';
    let cellVal = '';

    if (cell) {
      if (cell.f) {
        cellFormula = '=' + cell.f;
      } else if (cell.w !== undefined && cell.w !== null) {
        cellVal = cell.w;
      } else if (cell.v !== undefined && cell.v !== null) {
        cellVal = cell.v;
      }
    }

    cellFormulaEl.value = cellFormula ? cellFormula : String(cellVal !== undefined ? cellVal : '');

    calculateRangeStats();
  }

  gridBody.addEventListener('click', (e) => {
    const td = e.target.closest('td');
    if (td) {
      const r = parseInt(td.getAttribute('data-r'), 10);
      const c = parseInt(td.getAttribute('data-c'), 10);
      selectCell(r, c);
    }
  });

  function calculateRangeStats() {
    const col = selectedCell.col;
    let sum = 0;
    let numCount = 0;
    let min = Infinity;
    let max = -Infinity;

    if (activeSheetWs && activeSheetWs['!ref']) {
      const range = XLSX.utils.decode_range(activeSheetWs['!ref']);
      for (let r = 0; r <= range.e.r; r++) {
        const cellRef = XLSX.utils.encode_cell({ r, c: col });
        const cell = activeSheetWs[cellRef];
        if (cell && typeof cell.v === 'number') {
          const v = cell.v;
          sum += v;
          numCount++;
          if (v < min) min = v;
          if (v > max) max = v;
        }
      }
    }

    if (numCount > 0) {
      const avg = sum / numCount;
      statsSummaryEl.innerHTML = 'Col ' + colIndexToName(col) + ' — Sum: <strong>' + sum.toLocaleString() + '</strong> | Avg: <strong>' + avg.toFixed(2) + '</strong> | Min: <strong>' + min.toLocaleString() + '</strong> | Max: <strong>' + max.toLocaleString() + '</strong> (Count: ' + numCount + ')';
    } else {
      statsSummaryEl.textContent = '';
    }
  }

  // Live Search
  searchInput.addEventListener('input', () => {
    const term = searchInput.value.trim().toLowerCase();
    const cells = gridBody.querySelectorAll('td');
    let matchCount = 0;

    cells.forEach(cell => {
      cell.classList.remove('search-match');
      if (term && cell.textContent.toLowerCase().includes(term)) {
        cell.classList.add('search-match');
        matchCount++;
      }
    });

    searchCount.textContent = term ? (matchCount + ' found') : '';
  });

  // Zoom
  btnZoomIn.onclick = () => {
    if (currentZoom < 150) {
      currentZoom += 10;
      applyZoom();
    }
  };
  btnZoomOut.onclick = () => {
    if (currentZoom > 70) {
      currentZoom -= 10;
      applyZoom();
    }
  };
  function applyZoom() {
    zoomLevelEl.textContent = currentZoom + '%';
    document.getElementById('spreadsheetTable').style.transform = 'scale(' + (currentZoom / 100) + ')';
    document.getElementById('spreadsheetTable').style.transformOrigin = 'top left';
  }

  // Actions
  btnExportCsv.onclick = () => {
    if (!currentWorkbook || !activeSheetName) return;
    const ws = currentWorkbook.Sheets[activeSheetName];
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = (targetName.replace(/\.[^/.]+$/, '')) + '_' + activeSheetName + '.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  };

  btnPrint.onclick = () => window.print();

  btnDownloadOriginal.onclick = () => {
    if (originalRawBuffer) {
      const mimeType = isPdfBuffer(originalRawBuffer) ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      const blob = new Blob([originalRawBuffer], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = targetName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 10000);
    } else if (targetUrl) {
      chrome.runtime.sendMessage({
        action: 'BYPASS_AND_DOWNLOAD',
        url: targetUrl,
        filename: targetName
      });
    }
  };

  // Drag and drop file support
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      fileNameEl.textContent = file.name;
      updateFileBadge(file.name);
      const reader = new FileReader();
      reader.onload = (event) => {
        originalRawBuffer = event.target.result;
        renderWorkbook(event.target.result);
      };
      reader.readAsArrayBuffer(file);
    }
  });

  // About Modal Handlers
  const logoEl = document.querySelector('.logo');

  function openAboutModal() {
    if (aboutModal) {
      aboutModal.classList.remove('hidden');
      aboutModal.setAttribute('aria-hidden', 'false');
    }
  }

  function closeAboutModal() {
    if (aboutModal) {
      aboutModal.classList.add('hidden');
      aboutModal.setAttribute('aria-hidden', 'true');
    }
  }

  if (btnAbout) btnAbout.onclick = openAboutModal;
  if (logoEl) logoEl.onclick = openAboutModal;
  if (btnCloseAboutModal) btnCloseAboutModal.onclick = closeAboutModal;
  if (btnCloseAboutModalBtn) btnCloseAboutModalBtn.onclick = closeAboutModal;

  if (btnOpenAboutPage) {
    btnOpenAboutPage.onclick = () => {
      chrome.tabs.create({ url: chrome.runtime.getURL('about.html') });
    };
  }

  if (aboutModal) {
    aboutModal.onclick = (e) => {
      if (e.target === aboutModal) closeAboutModal();
    };
  }

  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && aboutModal && !aboutModal.classList.contains('hidden')) {
      closeAboutModal();
    }
  });

  loadSpreadsheet();
})();
