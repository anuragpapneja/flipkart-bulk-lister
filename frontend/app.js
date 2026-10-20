/**
 * Flipkart Bulk Lister — Clean UI Logic
 * Step 1: Upload → Step 2: Review & Generate → Step 3: Download
 */

document.addEventListener('DOMContentLoaded', () => {
  // ── DOM Refs ──
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');
  const btnBrowse = document.getElementById('btn-browse');
  const btnChangeFile = document.getElementById('btn-change-file');
  const btnGenerate = document.getElementById('btn-generate');
  const btnGenerateText = document.getElementById('btn-generate-text');
  const btnDownload = document.getElementById('btn-download');
  const btnDownloadText = document.getElementById('btn-download-text');
  const btnStartOver = document.getElementById('btn-start-over');

  // Panels
  const panelUpload = document.getElementById('panel-upload');
  const panelReview = document.getElementById('panel-review');
  const panelLoading = document.getElementById('panel-loading');
  const panelResults = document.getElementById('panel-results');

  // Step indicators
  const stepInd1 = document.getElementById('step-ind-1');
  const stepInd2 = document.getElementById('step-ind-2');
  const stepInd3 = document.getElementById('step-ind-3');
  const stepLine1 = document.getElementById('step-line-1');
  const stepLine2 = document.getElementById('step-line-2');

  // Review panel elements
  const fileName = document.getElementById('file-name');
  const fileMeta = document.getElementById('file-meta');
  const valRows = document.getElementById('val-rows');
  const valAttrs = document.getElementById('val-attrs');
  const valSheets = document.getElementById('val-sheets');

  // Results panel elements
  const successMsg = document.getElementById('success-msg');
  const statTotal = document.getElementById('stat-total');
  const previewTbody = document.getElementById('preview-tbody');
  const toast = document.getElementById('toast');

  // Settings modal
  const btnSettings = document.getElementById('btn-settings');
  const settingsModal = document.getElementById('settings-modal');
  const btnCloseModal = document.getElementById('btn-close-modal');
  const inputApiKey = document.getElementById('input-api-key');
  const btnToggleKey = document.getElementById('btn-toggle-key');
  const selectModel = document.getElementById('select-model');
  const inputBackendUrl = document.getElementById('input-backend-url');
  const btnTestBackend = document.getElementById('btn-test-backend');
  const btnSaveSettings = document.getElementById('btn-save-settings');

  // Header backend status
  const backendStatus = document.getElementById('backend-status');
  const backendStatusText = document.getElementById('backend-status-text');

  // ── State ──
  let currentFile = null;
  let currentAnalysis = null;
  let processedFileId = null;
  let processedFilename = null;

  // ── Helpers for Backend Resolution ──
  function getApiBaseUrl() {
    const saved = localStorage.getItem('flipkart_backend_url');
    if (saved !== null && saved.trim() !== '') {
      return saved.trim().replace(/\/+$/, '');
    }
    // If running on local server (FastAPI), use relative paths
    if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') {
      return '';
    }
    // If running on Firebase Hosting or external domain, default to local FastAPI server
    return 'http://127.0.0.1:8000';
  }

  async function checkBackendStatus() {
    const baseUrl = getApiBaseUrl();
    const testUrl = `${baseUrl}/api/health`;
    if (backendStatus) {
      backendStatus.className = 'status-pill checking';
      backendStatusText.textContent = 'Checking API...';
    }
    try {
      const res = await fetch(testUrl, { method: 'GET', mode: 'cors' });
      if (res.ok) {
        const data = await res.json();
        if (backendStatus) {
          backendStatus.className = 'status-pill online';
          backendStatusText.textContent = data.has_api_key ? 'API & Key Active' : 'API Connected';
          backendStatus.title = `Connected to ${baseUrl || 'Same-origin'}${data.has_api_key ? ' (Server .env Key Active)' : ''}`;
        }
        if (data.has_api_key && !inputApiKey.value) {
          inputApiKey.placeholder = 'Configured in server .env (Active)';
        }
        return true;
      } else {
        throw new Error(`HTTP ${res.status}`);
      }
    } catch (err) {
      if (backendStatus) {
        backendStatus.className = 'status-pill offline';
        backendStatusText.textContent = 'API Offline';
        backendStatus.title = `Cannot reach ${testUrl}. Make sure your local Python server is running!`;
      }
      return false;
    }
  }

  // ── Init Settings ──
  inputApiKey.value = localStorage.getItem('flipkart_gemini_api_key') || '';
  selectModel.value = localStorage.getItem('flipkart_gemini_model') || 'gemini-3.8-flash';
  if (inputBackendUrl) {
    inputBackendUrl.value = localStorage.getItem('flipkart_backend_url') || '';
  }

  // Check health on start & every 20 seconds
  checkBackendStatus();
  setInterval(checkBackendStatus, 20000);

  // ═══════════ HELPERS ═══════════

  function showToast(msg, duration = 3500) {
    toast.textContent = msg;
    toast.classList.remove('hidden');
    clearTimeout(toast._timer);
    toast._timer = setTimeout(() => toast.classList.add('hidden'), duration);
  }

  function escapeHtml(text) {
    if (!text) return '';
    return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function setStep(n) {
    // Reset all
    [stepInd1, stepInd2, stepInd3].forEach(el => { el.className = 'step'; });
    [stepLine1, stepLine2].forEach(el => { el.className = 'step-line'; });

    if (n >= 1) stepInd1.classList.add(n === 1 ? 'active' : 'done');
    if (n >= 2) { stepLine1.classList.add('done'); stepInd2.classList.add(n === 2 ? 'active' : 'done'); }
    if (n >= 3) { stepLine2.classList.add('done'); stepInd3.classList.add('active'); }
  }

  function showPanel(panel) {
    [panelUpload, panelReview, panelLoading, panelResults].forEach(p => p.classList.add('hidden'));
    panel.classList.remove('hidden');
  }

  // ═══════════ SETTINGS MODAL ═══════════

  btnSettings.addEventListener('click', () => {
    if (inputBackendUrl) {
      inputBackendUrl.value = localStorage.getItem('flipkart_backend_url') || '';
    }
    settingsModal.classList.remove('hidden');
  });
  if (backendStatus) {
    backendStatus.style.cursor = 'pointer';
    backendStatus.addEventListener('click', () => {
      if (inputBackendUrl) {
        inputBackendUrl.value = localStorage.getItem('flipkart_backend_url') || '';
      }
      settingsModal.classList.remove('hidden');
    });
  }
  btnCloseModal.addEventListener('click', () => settingsModal.classList.add('hidden'));
  settingsModal.addEventListener('click', (e) => {
    if (e.target === settingsModal) settingsModal.classList.add('hidden');
  });

  btnToggleKey.addEventListener('click', () => {
    if (inputApiKey.type === 'password') {
      inputApiKey.type = 'text';
      btnToggleKey.textContent = 'Hide';
    } else {
      inputApiKey.type = 'password';
      btnToggleKey.textContent = 'Show';
    }
  });

  if (btnTestBackend) {
    btnTestBackend.addEventListener('click', async () => {
      btnTestBackend.textContent = 'Testing...';
      const targetUrl = (inputBackendUrl.value.trim() || getApiBaseUrl()).replace(/\/+$/, '');
      try {
        const res = await fetch(`${targetUrl}/api/health`, { mode: 'cors' });
        if (res.ok) {
          showToast('✅ Backend connection successful!');
          if (backendStatus) {
            backendStatus.className = 'status-pill online';
            backendStatusText.textContent = 'API Connected';
          }
        } else {
          showToast(`⚠️ Server replied with status ${res.status}`);
        }
      } catch (e) {
        showToast('❌ Could not connect to backend server. Make sure run.py is active.');
      } finally {
        btnTestBackend.textContent = 'Test API';
      }
    });
  }

  btnSaveSettings.addEventListener('click', () => {
    localStorage.setItem('flipkart_gemini_api_key', inputApiKey.value.trim());
    localStorage.setItem('flipkart_gemini_model', selectModel.value);
    if (inputBackendUrl) {
      localStorage.setItem('flipkart_backend_url', inputBackendUrl.value.trim());
    }
    settingsModal.classList.add('hidden');
    showToast('Settings saved!');
    checkBackendStatus();
  });

  // ═══════════ STEP 1: UPLOAD ═══════════

  btnBrowse.addEventListener('click', () => fileInput.click());

  dropzone.addEventListener('click', (e) => {
    // Don't trigger if the user clicked the browse button (it already handles it)
    if (e.target.closest('#btn-browse')) return;
    fileInput.click();
  });

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('drag-active');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('drag-active');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-active');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files && fileInput.files.length > 0) {
      handleFile(fileInput.files[0]);
    }
  });

  // ── Analyze uploaded file ──
  async function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xlsm|xls)$/i)) {
      showToast('Please select a valid Excel file (.xls, .xlsx, or .xlsm)');
      return;
    }

    currentFile = file;
    showToast(`Analyzing ${file.name}...`);

    const formData = new FormData();
    formData.append('file', file);

    try {
      const resp = await fetch(`${getApiBaseUrl()}/api/analyze`, { method: 'POST', body: formData });
      const data = await resp.json();
      if (!resp.ok || !data.success) throw new Error(data.detail || 'Analysis failed');

      currentAnalysis = data.analysis;
      showReviewPanel(file, data.analysis);
    } catch (err) {
      if (err.message && err.message.toLowerCase().includes('fetch')) {
        showToast('⚠️ Backend unavailable. Click the status pill or ⚙️ to set your Render API URL.', 6000);
      } else {
        showToast('Error: ' + err.message);
      }
    }
  }

  // ═══════════ STEP 2: REVIEW ═══════════

  function showReviewPanel(file, analysis) {
    setStep(2);
    showPanel(panelReview);

    fileName.textContent = file.name;
    fileMeta.textContent = `Sheet: "${analysis.sheet_name}" • ${analysis.all_sheets.length} worksheets`;

    valRows.textContent = analysis.subsequent_active_rows_count;
    valAttrs.textContent = analysis.columns_detected.mirror_columns_count;
    valSheets.textContent = analysis.all_sheets.length;

    if (analysis.subsequent_active_rows_count === 0) {
      btnGenerate.disabled = true;
      btnGenerateText.textContent = 'No product rows found to fill';
      showToast('No rows with SKU or Image found after the 1st product.');
    } else {
      btnGenerate.disabled = false;
      btnGenerateText.textContent = `Generate ${analysis.subsequent_active_rows_count} Listings`;
    }
  }

  btnChangeFile.addEventListener('click', () => {
    currentFile = null;
    currentAnalysis = null;
    fileInput.value = '';
    setStep(1);
    showPanel(panelUpload);
  });

  // ═══════════ PROCESS ═══════════

  btnGenerate.addEventListener('click', async () => {
    if (!currentFile) return;

    setStep(2);
    showPanel(panelLoading);

    const apiKey = localStorage.getItem('flipkart_gemini_api_key') || '';
    const model = localStorage.getItem('flipkart_gemini_model') || 'gemini-3.8-flash';

    const formData = new FormData();
    formData.append('file', currentFile);
    formData.append('api_key', apiKey);
    formData.append('model', model);
    formData.append('preserve_existing', 'false'); // Mirror all 1st product attributes

    try {
      const resp = await fetch(`${getApiBaseUrl()}/api/process`, { method: 'POST', body: formData });
      const data = await resp.json();
      if (!resp.ok || !data.success) throw new Error(data.detail || 'Processing failed');

      processedFileId = data.file_id;
      processedFilename = data.filename;
      showResults(data.audit);
    } catch (err) {
      showPanel(panelReview);
      if (err.message && err.message.toLowerCase().includes('fetch')) {
        showToast('⚠️ Could not contact backend. Please check your Render URL in Settings ⚙️', 6000);
      } else {
        showToast('Error: ' + err.message);
      }
    }
  });

  // ═══════════ STEP 3: RESULTS ═══════════

  function showResults(audit) {
    setStep(3);
    showPanel(panelResults);

    const fname = processedFilename || 'Flipkart_Bulk_Listing.xls';

    // Stats
    statTotal.textContent = audit.total_processed_rows;
    successMsg.textContent = `${audit.total_processed_rows} product rows filled from 1st product data. File is ready to upload to Flipkart.`;

    // Download link
    const downloadUrl = `${getApiBaseUrl()}/api/download/${processedFileId}/${encodeURIComponent(fname)}`;
    btnDownload.href = downloadUrl;
    btnDownload.setAttribute('download', fname);
    btnDownloadText.textContent = `Download ${fname}`;

    // Preview table
    previewTbody.innerHTML = '';
    audit.processed_rows.forEach((row) => {
      const tr = document.createElement('tr');

      const kwParts = row.search_keywords ? row.search_keywords.split('::') : [];
      const kwHtml = kwParts.map(k => `<span class="kw-pill">${escapeHtml(k.trim())}</span>`).join('<span class="kw-sep">::</span>');

      tr.innerHTML = `
        <td><strong>#${row.row_index}</strong></td>
        <td><span class="sku-cell">${escapeHtml(row.sku)}</span></td>
        <td><span class="model-cell">${escapeHtml(row.model_name)}</span></td>
        <td><span class="char-tag">${row.model_name_len}/80</span></td>
        <td><div class="kw-container">${kwHtml}</div></td>
        <td>${row.main_image ? `<a href="${escapeHtml(row.main_image)}" target="_blank" rel="noopener" class="img-link" title="${escapeHtml(row.main_image)}">${escapeHtml(row.main_image)}</a>` : '<span style="color:var(--text-3)">—</span>'}</td>
        <td><span class="mirrored-badge">${row.mirrored_attributes_count} attrs</span></td>
      `;
      previewTbody.appendChild(tr);
    });

    panelResults.scrollIntoView({ behavior: 'smooth' });
    showToast(`Done! ${audit.total_processed_rows} listings generated. Click to download.`);
  }

  btnDownload.addEventListener('click', () => {
    showToast(`Downloading ${processedFilename || 'file'}...`);
  });

  // ── Start Over ──
  btnStartOver.addEventListener('click', () => {
    currentFile = null;
    currentAnalysis = null;
    processedFileId = null;
    processedFilename = null;
    fileInput.value = '';
    previewTbody.innerHTML = '';
    setStep(1);
    showPanel(panelUpload);
  });

  // ── Initial state ──
  setStep(1);
});
