document.addEventListener('DOMContentLoaded', () => {
  const interceptToggle = document.getElementById('interceptToggle');
  const fmtXlsx = document.getElementById('fmt_xlsx');
  const fmtXls = document.getElementById('fmt_xls');
  const fmtCsv = document.getElementById('fmt_csv');
  const fmtPdf = document.getElementById('fmt_pdf');

  chrome.storage.sync.get({
    interceptDownloads: true,
    supportedFormats: { xlsx: true, xls: true, csv: true, pdf: true }
  }, (items) => {
    interceptToggle.checked = items.interceptDownloads;
    fmtXlsx.checked = items.supportedFormats.xlsx !== false;
    fmtXls.checked = items.supportedFormats.xls !== false;
    fmtCsv.checked = items.supportedFormats.csv !== false;
    fmtPdf.checked = items.supportedFormats.pdf !== false;
  });

  function saveConfig() {
    chrome.storage.sync.set({
      interceptDownloads: interceptToggle.checked,
      supportedFormats: {
        xlsx: fmtXlsx.checked,
        xls: fmtXls.checked,
        csv: fmtCsv.checked,
        pdf: fmtPdf.checked
      }
    });
  }

  interceptToggle.addEventListener('change', saveConfig);
  fmtXlsx.addEventListener('change', saveConfig);
  fmtXls.addEventListener('change', saveConfig);
  fmtCsv.addEventListener('change', saveConfig);
  fmtPdf.addEventListener('change', saveConfig);

  const btnAboutPopup = document.getElementById('btnAboutPopup');
  if (btnAboutPopup) {
    btnAboutPopup.addEventListener('click', () => {
      chrome.tabs.create({ url: chrome.runtime.getURL('about.html') });
    });
  }
});
