/**
 * ExcelTab - Background Service Worker
 * Intercepts ONLY .xlsx, .xls, .csv, and .pdf file downloads and opens them in a viewer tab.
 */

const ALLOWED_EXTENSIONS = ['xlsx', 'xls', 'csv', 'pdf'];
const bypassedUrls = new Set();

function isSupportedExtension(filename, rawUrl) {
  const fn = (filename || '').toLowerCase();
  const url = (rawUrl || '').toLowerCase();

  return ALLOWED_EXTENSIONS.some(ext => {
    const dotExt = '.' + ext;
    return fn.endsWith(dotExt) || url.includes(dotExt + '?') || url.includes(dotExt + '#') || url.endsWith(dotExt);
  });
}

function generateFileId() {
  return 'sheet_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7);
}

function openViewerWithData(base64Data, fileName, openInForeground, callback) {
  const fileId = generateFileId();
  chrome.storage.local.set({
    [fileId]: {
      data: base64Data,
      name: fileName,
      timestamp: Date.now()
    }
  }, () => {
    const viewerUrl = chrome.runtime.getURL('viewer.html') +
      '?fileId=' + fileId +
      '&name=' + encodeURIComponent(fileName || 'Spreadsheet.xlsx');

    chrome.tabs.create({
      url: viewerUrl,
      active: openInForeground !== false
    }, (tab) => {
      if (callback) callback(tab);
    });
  });
}

// Listen to downloads before file is written to disk
chrome.downloads.onDeterminingFilename.addListener((downloadItem, suggest) => {
  // CRITICAL: If the download was initiated by ExcelTab itself (or from extension viewer tab), ALLOW IT directly!
  if (downloadItem.byExtensionId === chrome.runtime.id ||
      (downloadItem.url && downloadItem.url.includes(chrome.runtime.id)) ||
      (downloadItem.finalUrl && downloadItem.finalUrl.includes(chrome.runtime.id))) {
    suggest();
    return;
  }

  const rawUrl = (downloadItem.finalUrl || downloadItem.url || '').toLowerCase();
  if (bypassedUrls.has(rawUrl)) {
    bypassedUrls.delete(rawUrl);
    suggest();
    return;
  }

  const filename = (downloadItem.filename || rawUrl || '').toLowerCase();

  // STRICT CHECK: Reject non-supported formats (like .docx, .zip, .png, .pptx) immediately
  if (!isSupportedExtension(filename, rawUrl)) {
    suggest();
    return;
  }

  chrome.storage.sync.get({
    interceptDownloads: true,
    supportedFormats: { xlsx: true, xls: true, csv: true, pdf: true },
    openInForeground: true
  }, (config) => {
    if (!config.interceptDownloads) {
      suggest();
      return;
    }

    const isEnabledFormat = Object.keys(config.supportedFormats).some(ext => {
      return config.supportedFormats[ext] && (filename.endsWith('.' + ext) || rawUrl.includes('.' + ext));
    });

    if (isEnabledFormat) {
      // Intercept download: try fetching in page context first to preserve session
      if (downloadItem.tabId && downloadItem.tabId !== -1) {
        chrome.tabs.sendMessage(downloadItem.tabId, { action: 'FETCH_IN_PAGE', url: downloadItem.finalUrl || downloadItem.url }, (response) => {
          if (!chrome.runtime.lastError && response && response.success && response.data) {
            chrome.downloads.cancel(downloadItem.id, () => {
              chrome.downloads.erase({ id: downloadItem.id }, () => {});
              openViewerWithData(response.data, downloadItem.filename || 'Spreadsheet.xlsx', config.openInForeground);
            });
            return;
          }

          // If in-page fetch fails, cancel native disk download and open viewer with URL
          cancelAndOpenViewer(downloadItem, config.openInForeground);
        });
        return;
      }

      // Cancel native disk download and open viewer with URL
      cancelAndOpenViewer(downloadItem, config.openInForeground);
      return;
    }

    suggest();
  });

  return true; // async suggestion
});

function cancelAndOpenViewer(downloadItem, openInForeground) {
  chrome.downloads.cancel(downloadItem.id, () => {
    chrome.downloads.erase({ id: downloadItem.id }, () => {});

    const viewerUrl = chrome.runtime.getURL('viewer.html') +
      '?url=' + encodeURIComponent(downloadItem.finalUrl || downloadItem.url) +
      '&name=' + encodeURIComponent(downloadItem.filename || 'Spreadsheet.xlsx');

    chrome.tabs.create({
      url: viewerUrl,
      active: openInForeground !== false
    });
  });
}

function fetchWithSameOriginReferrer(url) {
  let refUrl = url;
  try {
    if (url.startsWith('http')) {
      const u = new URL(url);
      refUrl = u.origin + '/';
    }
  } catch (e) {}

  return fetch(url, {
    method: 'GET',
    credentials: 'include',
    referrer: refUrl,
    referrerPolicy: 'no-referrer-when-downgrade',
    headers: {
      'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel, text/csv, application/pdf, application/octet-stream, */*',
      'Accept-Language': 'en-US,en;q=0.9'
    }
  });
}

// Handle direct link clicks reported by content script, proxy fetches & bypass downloads
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'BYPASS_AND_DOWNLOAD') {
    if (message.url) {
      bypassedUrls.add(message.url.toLowerCase());
    }
    chrome.downloads.download({
      url: message.url,
      filename: message.filename || 'Spreadsheet.xlsx',
      saveAs: true
    }, (downloadId) => {
      sendResponse({ success: true, downloadId });
    });
    return true;
  }

  if (message.action === 'OPEN_EXCEL_VIEWER_WITH_DATA') {
    openViewerWithData(message.data, message.name, true, (tab) => {
      sendResponse({ success: true, tabId: tab ? tab.id : null });
    });
    return true;
  }

  if (message.action === 'OPEN_EXCEL_VIEWER') {
    const viewerUrl = chrome.runtime.getURL('viewer.html') +
      '?url=' + encodeURIComponent(message.url) +
      '&name=' + encodeURIComponent(message.name || 'Spreadsheet.xlsx');

    chrome.tabs.create({
      url: viewerUrl,
      active: true
    }, (newTab) => {
      sendResponse({ success: true, tabId: newTab.id });
    });
    return true;
  }

  if (message.action === 'FETCH_SPREADSHEET') {
    fetchWithSameOriginReferrer(message.url)
      .then(res => {
        if (!res.ok) throw new Error('HTTP ' + res.status + ' ' + res.statusText);
        return res.arrayBuffer();
      })
      .then(buffer => {
        const bytes = new Uint8Array(buffer);
        let binary = '';
        const chunkSize = 8192;
        for (let i = 0; i < bytes.length; i += chunkSize) {
          binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
        }
        const base64 = btoa(binary);
        sendResponse({ success: true, data: base64 });
      })
      .catch(err => {
        sendResponse({ success: false, error: err.message });
      });
    return true;
  }
});
