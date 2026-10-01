/**
 * ExcelTab - Content Script
 * Intercepts clicks on <a> tags targeting spreadsheet files.
 * Ignores file uploads, file input elements, and upload dropzones.
 */

(function() {
  const SPREADSHEET_REGEX = /\.(xlsx|xls|csv|pdf)($|[?#])/i;

  function isUploadElement(element, eventTarget) {
    if (!element) return false;

    // 1. Check if event target or element is a file input or inside a label for a file input
    if (eventTarget) {
      if (eventTarget.tagName === 'INPUT' && eventTarget.type === 'file') return true;
      const closestLabel = eventTarget.closest ? eventTarget.closest('label') : null;
      if (closestLabel) {
        if (closestLabel.querySelector('input[type="file"]')) return true;
        const labelFor = closestLabel.getAttribute('for');
        if (labelFor) {
          const linkedInput = document.getElementById(labelFor);
          if (linkedInput && linkedInput.type === 'file') return true;
        }
      }
    }

    // 2. Check if element itself contains an input[type="file"]
    if (element.querySelector && element.querySelector('input[type="file"]')) {
      return true;
    }

    // 3. Check if element is inside an upload container or form that has a file input
    const uploadContainer = element.closest('form[enctype*="multipart"], form[action*="upload"], [class*="upload"], [id*="upload"], [data-action*="upload"], [dropzone], [class*="dropzone"]');
    if (uploadContainer) {
      if (uploadContainer.querySelector('input[type="file"]')) {
        return true;
      }
    }

    // 4. Check if the element's href is non-downloadable (e.g. javascript:, #, or empty)
    const rawHref = element.getAttribute('href') || '';
    const trimmedHref = rawHref.trim();
    if (trimmedHref === '#' || trimmedHref.startsWith('javascript:') || trimmedHref === '') {
      return true; // Not a real download URL
    }

    // 5. Check type / role attributes for submit buttons or upload controls
    const typeAttr = element.getAttribute('type');
    const roleAttr = element.getAttribute('role');
    if (typeAttr === 'submit' || (roleAttr === 'button' && (element.className.toString().toLowerCase().includes('upload') || element.id.toLowerCase().includes('upload')))) {
      return true;
    }

    return false;
  }

  function isSpreadsheetLink(element) {
    if (!element || !element.href) return false;

    const href = element.href;
    const downloadAttr = element.getAttribute('download');

    // Ignore non-http/https/blob schemes
    if (!href.startsWith('http:') && !href.startsWith('https:') && !href.startsWith('blob:')) {
      return false;
    }

    // Check href extension
    if (SPREADSHEET_REGEX.test(href)) return true;

    // Check download attribute if specified
    if (downloadAttr && SPREADSHEET_REGEX.test(downloadAttr)) return true;

    return false;
  }

  function getCleanFileName(element) {
    const downloadAttr = element.getAttribute('download');
    if (downloadAttr && downloadAttr.trim()) return downloadAttr.trim();
    try {
      const urlObj = new URL(element.href);
      const pathname = urlObj.pathname;
      const lastPart = pathname.substring(pathname.lastIndexOf('/') + 1);
      return decodeURIComponent(lastPart) || 'Spreadsheet.xlsx';
    } catch {
      return 'Spreadsheet.xlsx';
    }
  }

  function fetchWithSameOriginReferrer(url) {
    let refUrl = window.location.href;
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
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel, text/csv, application/octet-stream, */*',
        'Accept-Language': 'en-US,en;q=0.9'
      }
    });
  }

  function arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    const chunkSize = 8192;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
  }

  function isHtmlBuffer(buffer, contentType) {
    if (contentType && contentType.toLowerCase().includes('text/html')) {
      return true;
    }
    const bytes = new Uint8Array(buffer);
    if (bytes.length < 4) return false;

    // Valid ZIP / XLSX starts with PK\x03\x04
    if (bytes[0] === 0x50 && bytes[1] === 0x4B && bytes[2] === 0x03 && bytes[3] === 0x04) {
      return false;
    }
    // Valid OLE2 / XLS starts with 0xD0 0xCF 0x11 0xE0
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

  function showInterceptorToast(filename) {
    const toast = document.createElement('div');
    toast.style.position = 'fixed';
    toast.style.bottom = '24px';
    toast.style.right = '24px';
    toast.style.zIndex = '2147483647';
    toast.style.background = '#FFFFFF';
    toast.style.color = '#0F172A';
    toast.style.padding = '12px 18px';
    toast.style.borderRadius = '8px';
    toast.style.boxShadow = '0 10px 25px -5px rgba(0, 0, 0, 0.15)';
    toast.style.fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    toast.style.fontSize = '13px';
    toast.style.display = 'flex';
    toast.style.alignItems = 'center';
    toast.style.gap = '10px';
    toast.style.border = '1px solid rgba(16, 185, 129, 0.5)';
    toast.style.transition = 'all 0.2s cubic-bezier(0.16, 1, 0.3, 1)';

    toast.innerHTML = `
      <span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#10B981;"></span>
      <span><strong>ExcelTab:</strong> Opening <em>${escapeHtml(filename)}</em> in browser tab...</span>
    `;

    document.body.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(8px)';
      setTimeout(() => toast.remove(), 250);
    }, 2200);
  }

  function escapeHtml(str) {
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // Intercept click events before native browser download takes over
  document.addEventListener('click', async function(e) {
    // If Shift or Alt key is held down, allow native download
    if (e.shiftKey || e.altKey) {
      return; // bypass to let user download if specifically intended
    }

    const target = e.target.closest('a');
    if (!target) return;

    // DO NOT INTERCEPT UPLOADS, file inputs, or upload dropzones!
    if (isUploadElement(target, e.target)) {
      return;
    }

    if (isSpreadsheetLink(target)) {
      const filename = getCleanFileName(target);
      const href = target.href;

      // If it's a blob: URL or explicit download attribute, handle in-page
      if (href.startsWith('blob:')) {
        e.preventDefault();
        e.stopPropagation();
        showInterceptorToast(filename);
        try {
          const response = await fetchWithSameOriginReferrer(href);
          if (!response.ok) throw new Error('HTTP ' + response.status);
          const buffer = await response.arrayBuffer();
          if (isHtmlBuffer(buffer, response.headers.get('content-type'))) {
            window.location.href = href; // fallback to browser native navigation
            return;
          }
          const base64Data = arrayBufferToBase64(buffer);
          chrome.runtime.sendMessage({
            action: 'OPEN_EXCEL_VIEWER_WITH_DATA',
            data: base64Data,
            name: filename
          });
        } catch (err) {
          console.warn('In-page blob fetch failed:', err);
          window.location.href = href;
        }
        return;
      }

      // For standard http/https links, perform a quick fetch check
      e.preventDefault();
      e.stopPropagation();
      showInterceptorToast(filename);

      try {
        const response = await fetchWithSameOriginReferrer(href);
        if (!response.ok) throw new Error('HTTP ' + response.status);
        const buffer = await response.arrayBuffer();
        const contentType = response.headers.get('content-type') || '';

        if (isHtmlBuffer(buffer, contentType)) {
          // Response is an HTML webpage, not a file! Trigger native browser navigation
          console.warn('URL returned HTML webpage instead of spreadsheet file. Triggering native navigation.');
          window.location.href = href;
          return;
        }

        const base64Data = arrayBufferToBase64(buffer);
        chrome.runtime.sendMessage({
          action: 'OPEN_EXCEL_VIEWER_WITH_DATA',
          data: base64Data,
          name: filename
        });
      } catch (err) {
        console.warn('In-page fetch failed, falling back to URL parameter:', err);
        chrome.runtime.sendMessage({
          action: 'OPEN_EXCEL_VIEWER',
          url: href,
          name: filename
        });
      }
    }
  }, true);

  // Listen for requests from background script to read blob: or page URLs
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if ((message.action === 'READ_BLOB_DATA' || message.action === 'FETCH_IN_PAGE') && message.url) {
      fetchWithSameOriginReferrer(message.url)
        .then(res => {
          if (!res.ok) throw new Error('HTTP ' + res.status);
          const contentType = res.headers.get('content-type') || '';
          return res.arrayBuffer().then(buffer => ({ buffer, contentType }));
        })
        .then(({ buffer, contentType }) => {
          if (isHtmlBuffer(buffer, contentType)) {
            sendResponse({ success: false, isHtml: true, error: 'Response is an HTML page' });
            return;
          }
          const base64 = arrayBufferToBase64(buffer);
          sendResponse({ success: true, data: base64 });
        })
        .catch(err => {
          console.error('Failed to read URL in content script:', err);
          sendResponse({ success: false, error: err.message });
        });
      return true; // async
    }
  });
})();
