/* ==========================================================================
   HD Converter Pro — Client Logic
   Handles: Tab switching | Word→PDF | PDF→Word
   ========================================================================== */

document.addEventListener('DOMContentLoaded', () => {

  /* -------------------------------------------------------------------------
     Tab Switching
     ---------------------------------------------------------------------- */
  const tabs = document.querySelectorAll('.tab-btn');
  const panels = {
    'word-to-pdf': document.getElementById('panel-word-to-pdf'),
    'pdf-to-word': document.getElementById('panel-pdf-to-word'),
    'merge-pdf':   document.getElementById('panel-merge-pdf'),
    'split-pdf':   document.getElementById('panel-split-pdf'),
    'image-to-pdf': document.getElementById('panel-image-to-pdf'),
  };

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const target = tab.dataset.tab;

      tabs.forEach(t => {
        t.classList.remove('tab-btn--active');
        t.setAttribute('aria-selected', 'false');
      });
      tab.classList.add('tab-btn--active');
      tab.setAttribute('aria-selected', 'true');

      Object.entries(panels).forEach(([key, panel]) => {
        panel.style.display = key === target ? 'block' : 'none';
      });
    });
  });

  /* -------------------------------------------------------------------------
     Format Selector (PDF → Word)
     ---------------------------------------------------------------------- */
  const formatOptions = document.querySelectorAll('.format-option');
  formatOptions.forEach(opt => {
    opt.addEventListener('click', () => {
      formatOptions.forEach(o => {
        o.classList.remove('format-option--active');
        const inp = o.querySelector('input');
        if (inp) inp.checked = false;
      });
      opt.classList.add('format-option--active');
      const input = opt.querySelector('input');
      if (input) input.checked = true;
    });
  });

  function getSelectedFormat() {
    const active = document.querySelector('.format-option--active input');
    return active ? active.value : 'docx';
  }

  /* =========================================================================
     Shared Utilities
     ====================================================================== */

  function formatBytes(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  /**
   * Create a generic file drop zone controller.
   * @param {object} cfg - Configuration object
   */
  function createDropZoneController(cfg) {
    const {
      dropZoneId,
      fileInputId,
      browseBtnId,
      filesSectionId,
      filesListId,
      fileCountId,
      clearAllBtnId,
      convertBtnId,
      convertBtnTextId,
      progressContainerId,
      progressBarFillId,
      progressStatusTextId,
      progressPercentageId,
      resultCardId,
      resultSubtitleId,
      resultListId,
      downloadZipBtnId,
      openFolderBtnId,
      acceptExtensions,   // e.g. ['.docx', '.doc']
      apiEndpoint,        // e.g. '/api/convert'
      extraFormDataFn,    // () => FormData extras (optional)
      renderResultItemFn, // (item) => HTMLElement
      defaultBtnLabel,
      convertingLabel,
      successLabel,
    } = cfg;

    const dropZone      = document.getElementById(dropZoneId);
    const fileInput     = document.getElementById(fileInputId);
    const browseBtn     = document.getElementById(browseBtnId);
    const filesSection  = document.getElementById(filesSectionId);
    const filesList     = document.getElementById(filesListId);
    const fileCountSpan = document.getElementById(fileCountId);
    const clearAllBtn   = document.getElementById(clearAllBtnId);
    const convertBtn    = document.getElementById(convertBtnId);
    const convertBtnTxt = document.getElementById(convertBtnTextId);
    const progressCont  = document.getElementById(progressContainerId);
    const progressFill  = document.getElementById(progressBarFillId);
    const progressText  = document.getElementById(progressStatusTextId);
    const progressPct   = document.getElementById(progressPercentageId);
    const resultCard    = document.getElementById(resultCardId);
    const resultSubtitle= document.getElementById(resultSubtitleId);
    const resultList    = document.getElementById(resultListId);
    const zipBtn        = document.getElementById(downloadZipBtnId);
    const folderBtn     = document.getElementById(openFolderBtnId);

    let selectedFiles = [];

    // ── Drag & Drop ─────────────────────────────────────────────────────────
    ['dragenter', 'dragover'].forEach(evt => {
      dropZone.addEventListener(evt, e => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(evt => {
      dropZone.addEventListener(evt, e => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('dragover');
      });
    });

    dropZone.addEventListener('drop', e => {
      handleNewFiles(Array.from(e.dataTransfer.files));
    });

    dropZone.addEventListener('click', e => {
      if (e.target !== browseBtn && !browseBtn.contains(e.target)) {
        fileInput.click();
      }
    });

    browseBtn.addEventListener('click', e => {
      e.stopPropagation();
      fileInput.click();
    });

    fileInput.addEventListener('change', e => {
      handleNewFiles(Array.from(e.target.files));
      fileInput.value = '';
    });

    clearAllBtn.addEventListener('click', () => {
      selectedFiles = [];
      updateUI();
    });

    // ── File Handling ────────────────────────────────────────────────────────
    function handleNewFiles(files) {
      const added = [];
      files.forEach(file => {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (!acceptExtensions.includes(ext)) return;
        const dup = selectedFiles.some(f => f.name === file.name && f.size === file.size);
        if (!dup) { selectedFiles.push(file); added.push(file); }
      });

      if (added.length === 0 && files.length > 0) {
        alert(`Harap masukkan file ${acceptExtensions.join(' atau ')} saja.`);
      }
      updateUI();
    }

    function updateUI() {
      filesList.innerHTML = '';
      fileCountSpan.textContent = selectedFiles.length;

      if (selectedFiles.length > 0) {
        filesSection.style.display = 'block';
        convertBtn.disabled = false;
        convertBtnTxt.textContent = `${defaultBtnLabel} (${selectedFiles.length} File)`;

        selectedFiles.forEach((file, index) => {
          const item = document.createElement('div');
          item.className = 'file-item';

          const extIcon = acceptExtensions[0] === '.pdf'
            ? `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
              </svg>`
            : `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
              </svg>`;

          item.innerHTML = `
            <div class="file-info">
              <div class="file-icon">${extIcon}</div>
              <div class="file-text">
                <div class="file-name" title="${file.name}">${file.name}</div>
                <div class="file-meta">${formatBytes(file.size)} &bull; Siap Konversi</div>
              </div>
            </div>
            <button type="button" class="remove-file-btn" data-index="${index}" title="Hapus file ini">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <line x1="18" y1="6" x2="6" y2="18"></line>
                <line x1="6" y1="6" x2="18" y2="18"></line>
              </svg>
            </button>
          `;

          item.querySelector('.remove-file-btn').addEventListener('click', e => {
            e.stopPropagation();
            selectedFiles.splice(index, 1);
            updateUI();
          });

          filesList.appendChild(item);
        });
      } else {
        filesSection.style.display = 'none';
        convertBtn.disabled = true;
        convertBtnTxt.textContent = defaultBtnLabel;
      }
    }

    // ── Progress Helpers ─────────────────────────────────────────────────────
    function setProgress(pct, text) {
      progressFill.style.width = `${pct}%`;
      progressPct.textContent = `${pct}%`;
      if (text) progressText.textContent = text;
    }

    // ── Convert ──────────────────────────────────────────────────────────────
    convertBtn.addEventListener('click', async () => {
      if (selectedFiles.length === 0) return;

      resultCard.style.display = 'none';
      progressCont.style.display = 'flex';
      convertBtn.disabled = true;
      convertBtnTxt.textContent = convertingLabel;
      setProgress(10, 'Mengunggah dan menyiapkan engine konversi...');

      const formData = new FormData();
      selectedFiles.forEach(f => formData.append('files', f));

      // Append extra fields (e.g. options for Word→PDF)
      if (typeof extraFormDataFn === 'function') {
        extraFormDataFn(formData);
      }

      // Simulated progress ticker
      let pct = 10;
      const ticker = setInterval(() => {
        if (pct < 85) {
          pct += Math.floor(Math.random() * 7) + 2;
          if (pct > 85) pct = 85;
          let msg = 'Memproses dokumen dengan Word Engine...';
          if (pct > 40 && pct < 70) msg = 'Engine sedang merender dan menyusun konten...';
          else if (pct >= 70) msg = 'Menyempurnakan dan menyimpan file hasil...';
          setProgress(pct, msg);
        }
      }, 500);

      try {
        const response = await fetch(apiEndpoint, { method: 'POST', body: formData });
        clearInterval(ticker);
        const data = await response.json();

        if (!response.ok || !data.success) {
          throw new Error(
            data.error ||
            (data.errors && data.errors[0]?.error) ||
            'Konversi gagal — tidak ada file berhasil diproses.'
          );
        }

        setProgress(100, successLabel);

        setTimeout(() => {
          progressCont.style.display = 'none';
          renderResults(data);
          convertBtn.disabled = false;
          convertBtnTxt.textContent = `${defaultBtnLabel} (${selectedFiles.length} File)`;
        }, 400);

      } catch (err) {
        clearInterval(ticker);
        progressCont.style.display = 'none';
        convertBtn.disabled = false;
        convertBtnTxt.textContent = `${defaultBtnLabel} (${selectedFiles.length} File)`;
        alert('Terjadi kesalahan:\n' + err.message);
      }
    });

    // ── Render Results ───────────────────────────────────────────────────────
    function renderResults(data) {
      resultCard.style.display = 'block';
      resultList.innerHTML = '';

      const count = data.converted.length;
      resultSubtitle.textContent = `Berhasil mengonversi ${count} file.`;

      if (data.zip_url) {
        zipBtn.href = data.zip_url;
        zipBtn.style.display = 'inline-flex';
      } else {
        zipBtn.style.display = 'none';
      }

      data.converted.forEach(item => {
        const el = renderResultItemFn(item);
        resultList.appendChild(el);
      });

      resultCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    // ── Open Folder ──────────────────────────────────────────────────────────
    if (folderBtn) {
      folderBtn.addEventListener('click', async () => {
        try {
          const res = await fetch('/api/open-folder', { method: 'POST' });
          const d = await res.json();
          if (!d.success) alert('Gagal membuka folder: ' + d.error);
        } catch (e) {
          alert('Error: ' + e.message);
        }
      });
    }
  }

  /* =========================================================================
     Word → PDF Controller
     ====================================================================== */
  createDropZoneController({
    dropZoneId:          'drop-zone-w2p',
    fileInputId:         'file-input-w2p',
    browseBtnId:         'browse-btn-w2p',
    filesSectionId:      'selected-files-section-w2p',
    filesListId:         'files-list-w2p',
    fileCountId:         'file-count-w2p',
    clearAllBtnId:       'clear-all-btn-w2p',
    convertBtnId:        'convert-btn-w2p',
    convertBtnTextId:    'convert-btn-text-w2p',
    progressContainerId: 'progress-container-w2p',
    progressBarFillId:   'progress-bar-fill-w2p',
    progressStatusTextId:'progress-status-text-w2p',
    progressPercentageId:'progress-percentage-w2p',
    resultCardId:        'results-card-w2p',
    resultSubtitleId:    'results-subtitle-w2p',
    resultListId:        'converted-list-w2p',
    downloadZipBtnId:    'download-zip-btn-w2p',
    openFolderBtnId:     'open-folder-btn',
    acceptExtensions:    ['.docx', '.doc'],
    apiEndpoint:         '/api/convert',
    defaultBtnLabel:     'Mulai Konversi ke HD PDF',
    convertingLabel:     'Sedang Mengonversi ke PDF...',
    successLabel:        'Konversi ke PDF selesai!',

    extraFormDataFn: (fd) => {
      fd.append('optimize_print',  document.getElementById('opt-optimize-print').checked  ? 'true' : 'false');
      fd.append('preserve_fonts',  document.getElementById('opt-preserve-fonts').checked  ? 'true' : 'false');
      fd.append('create_bookmarks',document.getElementById('opt-bookmarks').checked        ? 'true' : 'false');
    },

    renderResultItemFn: (item) => {
      let imgMeta = '';
      if (item.image_info && item.image_info.image_count > 0) {
        imgMeta = `&bull; ${item.image_info.image_count} Gambar (${item.image_info.estimated_quality})`;
      }
      const div = document.createElement('div');
      div.className = 'converted-item';
      div.innerHTML = `
        <div class="converted-info">
          <div class="pdf-icon">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
              <polyline points="14 2 14 8 20 8"></polyline>
              <line x1="16" y1="13" x2="8" y2="13"></line>
              <line x1="16" y1="17" x2="8" y2="17"></line>
            </svg>
          </div>
          <div>
            <div class="converted-name">${item.pdf_name}</div>
            <div class="converted-details">
              <span>${item.size_kb} KB</span>
              <span class="tag-hd">Ultra HD 300 DPI</span>
              ${imgMeta ? `<span>${imgMeta}</span>` : ''}
            </div>
          </div>
        </div>
        <a href="${item.download_url}" class="btn btn-primary" download>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
            <polyline points="7 10 12 15 17 10"></polyline>
            <line x1="12" y1="15" x2="12" y2="3"></line>
          </svg>
          Unduh PDF
        </a>
      `;
      return div;
    },
  });

  /* =========================================================================
     PDF → Word Controller
     ====================================================================== */
  createDropZoneController({
    dropZoneId:          'drop-zone-p2w',
    fileInputId:         'file-input-p2w',
    browseBtnId:         'browse-btn-p2w',
    filesSectionId:      'selected-files-section-p2w',
    filesListId:         'files-list-p2w',
    fileCountId:         'file-count-p2w',
    clearAllBtnId:       'clear-all-btn-p2w',
    convertBtnId:        'convert-btn-p2w',
    convertBtnTextId:    'convert-btn-text-p2w',
    progressContainerId: 'progress-container-p2w',
    progressBarFillId:   'progress-bar-fill-p2w',
    progressStatusTextId:'progress-status-text-p2w',
    progressPercentageId:'progress-percentage-p2w',
    resultCardId:        'results-card-p2w',
    resultSubtitleId:    'results-subtitle-p2w',
    resultListId:        'converted-list-p2w',
    downloadZipBtnId:    'download-zip-btn-p2w',
    openFolderBtnId:     'open-folder-btn-p2w',
    acceptExtensions:    ['.pdf'],
    apiEndpoint:         '/api/pdf-to-word',
    defaultBtnLabel:     'Konversi ke Word',
    convertingLabel:     'Word Engine membuka PDF...',
    successLabel:        'Konversi selesai!',

    extraFormDataFn: (fd) => {
      fd.append('output_format', getSelectedFormat());
    },

    renderResultItemFn: (item) => {
      const ext = item.format || 'docx';
      const div = document.createElement('div');
      div.className = 'converted-item';
      div.innerHTML = `
        <div class="converted-info">
          <div class="file-icon">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
              <polyline points="14 2 14 8 20 8"></polyline>
              <line x1="16" y1="13" x2="8" y2="13"></line>
              <line x1="16" y1="17" x2="8" y2="17"></line>
            </svg>
          </div>
          <div>
            <div class="converted-name">${item.word_name}</div>
            <div class="converted-details">
              <span>${item.size_kb} KB</span>
              <span class="tag-hd">.${ext} &bull; Siap Edit</span>
            </div>
          </div>
        </div>
        <a href="${item.download_url}" class="btn btn-primary" download>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
            <polyline points="7 10 12 15 17 10"></polyline>
            <line x1="12" y1="15" x2="12" y2="3"></line>
          </svg>
          Unduh .${ext}
        </a>
      `;
      return div;
    },
  });

  /* =========================================================================
     PDF Merge Controller
     ====================================================================== */
  (() => {
    const dropZone      = document.getElementById('drop-zone-merge');
    const fileInput     = document.getElementById('file-input-merge');
    const browseBtn     = document.getElementById('browse-btn-merge');
    const filesSection  = document.getElementById('selected-files-section-merge');
    const filesList     = document.getElementById('files-list-merge');
    const fileCountSpan = document.getElementById('file-count-merge');
    const clearAllBtn   = document.getElementById('clear-all-btn-merge');
    const convertBtn    = document.getElementById('convert-btn-merge');
    const convertBtnTxt = document.getElementById('convert-btn-text-merge');
    const outputNameInp = document.getElementById('output-name-merge');
    const progressCont  = document.getElementById('progress-container-merge');
    const progressFill  = document.getElementById('progress-bar-fill-merge');
    const progressText  = document.getElementById('progress-status-text-merge');
    const progressPct   = document.getElementById('progress-percentage-merge');
    const resultCard    = document.getElementById('results-card-merge');
    const resultSubtitle= document.getElementById('results-subtitle-merge');
    const downloadBtn   = document.getElementById('download-btn-merge');
    const folderBtn     = document.getElementById('open-folder-btn-merge');

    let selectedFiles = [];

    ['dragenter', 'dragover'].forEach(evt => {
      dropZone.addEventListener(evt, e => {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.add('dragover');
      });
    });
    ['dragleave', 'drop'].forEach(evt => {
      dropZone.addEventListener(evt, e => {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.remove('dragover');
      });
    });
    dropZone.addEventListener('drop', e => {
      handleFiles(Array.from(e.dataTransfer.files));
    });
    dropZone.addEventListener('click', e => {
      if (e.target !== browseBtn && !browseBtn.contains(e.target)) fileInput.click();
    });
    browseBtn.addEventListener('click', e => { e.stopPropagation(); fileInput.click(); });
    fileInput.addEventListener('change', e => {
      handleFiles(Array.from(e.target.files));
      fileInput.value = '';
    });
    clearAllBtn.addEventListener('click', () => { selectedFiles = []; updateUI(); });

    function handleFiles(files) {
      files.forEach(f => {
        if (f.name.toLowerCase().endsWith('.pdf')) {
          if (!selectedFiles.some(cur => cur.name === f.name && cur.size === f.size)) {
            selectedFiles.push(f);
          }
        }
      });
      updateUI();
    }

    function moveItem(fromIdx, toIdx) {
      if (toIdx < 0 || toIdx >= selectedFiles.length) return;
      const item = selectedFiles.splice(fromIdx, 1)[0];
      selectedFiles.splice(toIdx, 0, item);
      updateUI();
    }

    function updateUI() {
      filesList.innerHTML = '';
      fileCountSpan.textContent = selectedFiles.length;

      if (selectedFiles.length > 0) {
        filesSection.style.display = 'block';
        convertBtn.disabled = selectedFiles.length < 2;
        convertBtnTxt.textContent = selectedFiles.length < 2
          ? 'Pilih Minimal 2 File PDF'
          : `Gabungkan ${selectedFiles.length} PDF`;

        selectedFiles.forEach((file, index) => {
          const item = document.createElement('div');
          item.className = 'file-item';
          item.innerHTML = `
            <div class="file-info">
              <div class="file-icon" style="color: #10b981; background: rgba(16, 185, 129, 0.12); border-color: rgba(16, 185, 129, 0.3);">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                  <polyline points="14 2 14 8 20 8"></polyline>
                </svg>
              </div>
              <div class="file-text">
                <div class="file-name" title="${file.name}">${index + 1}. ${file.name}</div>
                <div class="file-meta">${formatBytes(file.size)}</div>
              </div>
            </div>
            <div class="file-actions-group">
              <div class="order-btn-group">
                <button type="button" class="order-btn btn-up" title="Geser ke Atas" ${index === 0 ? 'disabled' : ''}>▲</button>
                <button type="button" class="order-btn btn-down" title="Geser ke Bawah" ${index === selectedFiles.length - 1 ? 'disabled' : ''}>▼</button>
              </div>
              <button type="button" class="remove-file-btn" title="Hapus file ini">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line>
                </svg>
              </button>
            </div>
          `;

          item.querySelector('.btn-up').addEventListener('click', (e) => { e.stopPropagation(); moveItem(index, index - 1); });
          item.querySelector('.btn-down').addEventListener('click', (e) => { e.stopPropagation(); moveItem(index, index + 1); });
          item.querySelector('.remove-file-btn').addEventListener('click', (e) => {
            e.stopPropagation();
            selectedFiles.splice(index, 1);
            updateUI();
          });

          filesList.appendChild(item);
        });
      } else {
        filesSection.style.display = 'none';
        convertBtn.disabled = true;
        convertBtnTxt.textContent = 'Gabungkan Semua PDF';
      }
    }

    convertBtn.addEventListener('click', async () => {
      if (selectedFiles.length < 2) return;

      resultCard.style.display = 'none';
      progressCont.style.display = 'flex';
      convertBtn.disabled = true;
      convertBtnTxt.textContent = 'Sedang Menggabungkan...';
      progressFill.style.width = '35%';
      progressPct.textContent = '35%';
      progressText.textContent = 'Membaca dan menggabungkan halaman berkas PDF...';

      const fd = new FormData();
      selectedFiles.forEach(f => fd.append('files', f));
      if (outputNameInp.value.trim()) {
        fd.append('output_name', outputNameInp.value.trim());
      }

      try {
        const res = await fetch('/api/pdf/merge', { method: 'POST', body: fd });
        const data = await res.json();
        if (!res.ok || !data.success) {
          throw new Error(data.error || 'Gagal menggabungkan PDF');
        }

        progressFill.style.width = '100%';
        progressPct.textContent = '100%';
        progressText.textContent = 'Selesai digabungkan!';

        setTimeout(() => {
          progressCont.style.display = 'none';
          resultCard.style.display = 'block';
          resultSubtitle.textContent = `Berhasil menggabungkan ${data.total_files_merged} berkas (${data.size_kb} KB).`;
          downloadBtn.href = data.download_url;
          downloadBtn.setAttribute('download', data.filename);
          convertBtn.disabled = false;
          convertBtnTxt.textContent = `Gabungkan ${selectedFiles.length} PDF`;
          resultCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }, 300);

      } catch (err) {
        progressCont.style.display = 'none';
        convertBtn.disabled = false;
        convertBtnTxt.textContent = `Gabungkan ${selectedFiles.length} PDF`;
        alert('Terjadi kesalahan saat menggabungkan PDF:\n' + err.message);
      }
    });

    if (folderBtn) {
      folderBtn.addEventListener('click', async () => {
        try { await fetch('/api/open-folder', { method: 'POST' }); } catch (e) { alert(e.message); }
      });
    }
  })();


  /* =========================================================================
     PDF Split Controller
     ====================================================================== */
  (() => {
    const dropZone         = document.getElementById('drop-zone-split');
    const fileInput        = document.getElementById('file-input-split');
    const browseBtn        = document.getElementById('browse-btn-split');
    const filesSection     = document.getElementById('selected-files-section-split');
    const filenameDisplay  = document.getElementById('split-filename-display');
    const filemetaDisplay  = document.getElementById('split-filemeta-display');
    const pagesBadge       = document.getElementById('split-pages-badge');
    const pagesCountText   = document.getElementById('split-pages-count-text');
    const clearBtn         = document.getElementById('clear-btn-split');
    const optionsBar       = document.getElementById('split-options-bar');
    const rangeInput       = document.getElementById('split-range-input');
    const rangeWrapper     = document.getElementById('split-range-input-wrapper');
    const modeRangeLabel   = document.getElementById('split-mode-range-label');
    const modeAllLabel     = document.getElementById('split-mode-all-label');
    const convertBtn       = document.getElementById('convert-btn-split');
    const convertBtnTxt    = document.getElementById('convert-btn-text-split');
    const progressCont     = document.getElementById('progress-container-split');
    const progressFill     = document.getElementById('progress-bar-fill-split');
    const progressText     = document.getElementById('progress-status-text-split');
    const progressPct      = document.getElementById('progress-percentage-split');
    const resultCard       = document.getElementById('results-card-split');
    const resultSubtitle   = document.getElementById('results-subtitle-split');
    const downloadBtn      = document.getElementById('download-btn-split');
    const downloadBtnTxt   = document.getElementById('download-btn-text-split');
    const folderBtn        = document.getElementById('open-folder-btn-split');

    let currentFile = null;
    let totalPdfPages = 0;
    let splitMode = 'range';

    // Mode Selector toggle
    modeRangeLabel.addEventListener('click', () => {
      splitMode = 'range';
      modeRangeLabel.classList.add('format-option--active');
      modeAllLabel.classList.remove('format-option--active');
      rangeWrapper.style.display = 'block';
    });

    modeAllLabel.addEventListener('click', () => {
      splitMode = 'all';
      modeAllLabel.classList.add('format-option--active');
      modeRangeLabel.classList.remove('format-option--active');
      rangeWrapper.style.display = 'none';
    });

    ['dragenter', 'dragover'].forEach(evt => {
      dropZone.addEventListener(evt, e => {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.add('dragover');
      });
    });
    ['dragleave', 'drop'].forEach(evt => {
      dropZone.addEventListener(evt, e => {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.remove('dragover');
      });
    });
    dropZone.addEventListener('drop', e => {
      if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
    });
    dropZone.addEventListener('click', e => {
      if (e.target !== browseBtn && !browseBtn.contains(e.target)) fileInput.click();
    });
    browseBtn.addEventListener('click', e => { e.stopPropagation(); fileInput.click(); });
    fileInput.addEventListener('change', e => {
      if (e.target.files.length > 0) handleFile(e.target.files[0]);
      fileInput.value = '';
    });

    clearBtn.addEventListener('click', () => {
      currentFile = null;
      totalPdfPages = 0;
      filesSection.style.display = 'none';
      optionsBar.style.display = 'none';
      dropZone.style.display = 'block';
      convertBtn.disabled = true;
      rangeInput.value = '';
    });

    async function handleFile(file) {
      if (!file.name.toLowerCase().endsWith('.pdf')) {
        alert('Harap pilih file PDF (.pdf).');
        return;
      }

      currentFile = file;
      filenameDisplay.textContent = file.name;
      filemetaDisplay.textContent = `${formatBytes(file.size)} • Membaca info...`;
      dropZone.style.display = 'none';
      filesSection.style.display = 'block';
      optionsBar.style.display = 'grid';
      convertBtn.disabled = true;

      // Auto inspect PDF info
      try {
        const fd = new FormData();
        fd.append('file', file);
        const res = await fetch('/api/pdf/info', { method: 'POST', body: fd });
        const data = await res.json();
        if (res.ok && data.success) {
          totalPdfPages = data.info.total_pages;
          pagesCountText.textContent = `${totalPdfPages} Halaman`;
          filemetaDisplay.textContent = `${formatBytes(file.size)} • Siap Dipisah`;
          if (totalPdfPages > 1) {
            rangeInput.placeholder = `Misal: 1-${Math.min(totalPdfPages, 3)}, ${totalPdfPages}`;
          } else {
            rangeInput.placeholder = `1`;
          }
          convertBtn.disabled = false;
        } else {
          throw new Error(data.error || 'Gagal membaca info PDF');
        }
      } catch (err) {
        filemetaDisplay.textContent = `${formatBytes(file.size)} • ${err.message}`;
        convertBtn.disabled = false;
      }
    }

    convertBtn.addEventListener('click', async () => {
      if (!currentFile) return;

      if (splitMode === 'range' && !rangeInput.value.trim()) {
        alert('Harap masukkan rentang halaman yang ingin diekstrak (contoh: 1-2 atau 1, 3).');
        rangeInput.focus();
        return;
      }

      resultCard.style.display = 'none';
      progressCont.style.display = 'flex';
      convertBtn.disabled = true;
      convertBtnTxt.textContent = 'Memisahkan Halaman...';
      progressFill.style.width = '40%';
      progressPct.textContent = '40%';
      progressText.textContent = 'Mengekstrak halaman dokumen PDF...';

      const fd = new FormData();
      fd.append('file', currentFile);
      fd.append('mode', splitMode);
      if (splitMode === 'range') {
        fd.append('page_range', rangeInput.value.trim());
      }

      try {
        const res = await fetch('/api/pdf/split', { method: 'POST', body: fd });
        const data = await res.json();
        if (!res.ok || !data.success) {
          throw new Error(data.error || 'Gagal memisahkan PDF');
        }

        progressFill.style.width = '100%';
        progressPct.textContent = '100%';
        progressText.textContent = 'Pemisahan halaman selesai!';

        setTimeout(() => {
          progressCont.style.display = 'none';
          resultCard.style.display = 'block';
          if (data.is_zip) {
            resultSubtitle.textContent = `Berhasil memisahkan ${data.extracted_pages_count} halaman ke file ZIP (${data.size_kb} KB).`;
            downloadBtnTxt.textContent = 'Unduh Berkas ZIP';
          } else {
            resultSubtitle.textContent = `Berhasil mengekstrak ${data.extracted_pages_count} halaman terpilih (${data.size_kb} KB).`;
            downloadBtnTxt.textContent = 'Unduh PDF Hasil';
          }
          downloadBtn.href = data.download_url;
          downloadBtn.setAttribute('download', data.filename);
          convertBtn.disabled = false;
          convertBtnTxt.textContent = 'Mulai Pisahkan PDF';
          resultCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }, 300);

      } catch (err) {
        progressCont.style.display = 'none';
        convertBtn.disabled = false;
        convertBtnTxt.textContent = 'Mulai Pisahkan PDF';
        alert('Terjadi kesalahan saat memisahkan PDF:\n' + err.message);
      }
    });

    if (folderBtn) {
      folderBtn.addEventListener('click', async () => {
        try { await fetch('/api/open-folder', { method: 'POST' }); } catch (e) { alert(e.message); }
      });
    }
  })();


  /* =========================================================================
     Image to PDF Controller
     ====================================================================== */
  (() => {
    const dropZone         = document.getElementById('drop-zone-img');
    const fileInput        = document.getElementById('file-input-img');
    const browseBtn        = document.getElementById('browse-btn-img');
    const filesSection     = document.getElementById('selected-files-section-img');
    const filesList        = document.getElementById('files-list-img');
    const fileCountSpan    = document.getElementById('file-count-img');
    const clearAllBtn      = document.getElementById('clear-all-btn-img');
    const outputNameInp    = document.getElementById('output-name-img');
    const fitLabel         = document.getElementById('pagesize-fit-label');
    const a4Label          = document.getElementById('pagesize-a4-label');
    const convertBtn       = document.getElementById('convert-btn-img');
    const convertBtnTxt    = document.getElementById('convert-btn-text-img');
    const progressCont     = document.getElementById('progress-container-img');
    const progressFill     = document.getElementById('progress-bar-fill-img');
    const progressText     = document.getElementById('progress-status-text-img');
    const progressPct      = document.getElementById('progress-percentage-img');
    const resultCard       = document.getElementById('results-card-img');
    const resultSubtitle   = document.getElementById('results-subtitle-img');
    const downloadBtn      = document.getElementById('download-btn-img');
    const folderBtn        = document.getElementById('open-folder-btn-img');

    let selectedImages = [];
    let pageSize = 'fit';

    fitLabel.addEventListener('click', () => {
      pageSize = 'fit';
      fitLabel.classList.add('format-option--active');
      a4Label.classList.remove('format-option--active');
    });

    a4Label.addEventListener('click', () => {
      pageSize = 'a4';
      a4Label.classList.add('format-option--active');
      fitLabel.classList.remove('format-option--active');
    });

    ['dragenter', 'dragover'].forEach(evt => {
      dropZone.addEventListener(evt, e => {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.add('dragover');
      });
    });
    ['dragleave', 'drop'].forEach(evt => {
      dropZone.addEventListener(evt, e => {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.remove('dragover');
      });
    });
    dropZone.addEventListener('drop', e => {
      handleImages(Array.from(e.dataTransfer.files));
    });
    dropZone.addEventListener('click', e => {
      if (e.target !== browseBtn && !browseBtn.contains(e.target)) fileInput.click();
    });
    browseBtn.addEventListener('click', e => { e.stopPropagation(); fileInput.click(); });
    fileInput.addEventListener('change', e => {
      handleImages(Array.from(e.target.files));
      fileInput.value = '';
    });
    clearAllBtn.addEventListener('click', () => { selectedImages = []; updateUI(); });

    const validImgExts = ['.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tiff'];
    function handleImages(files) {
      files.forEach(f => {
        const ext = '.' + f.name.split('.').pop().toLowerCase();
        if (validImgExts.includes(ext)) {
          if (!selectedImages.some(cur => cur.name === f.name && cur.size === f.size)) {
            selectedImages.push(f);
          }
        }
      });
      updateUI();
    }

    function moveItem(fromIdx, toIdx) {
      if (toIdx < 0 || toIdx >= selectedImages.length) return;
      const item = selectedImages.splice(fromIdx, 1)[0];
      selectedImages.splice(toIdx, 0, item);
      updateUI();
    }

    function updateUI() {
      filesList.innerHTML = '';
      fileCountSpan.textContent = selectedImages.length;

      if (selectedImages.length > 0) {
        filesSection.style.display = 'block';
        convertBtn.disabled = false;
        convertBtnTxt.textContent = `Konversi ${selectedImages.length} Gambar ke PDF`;

        selectedImages.forEach((file, index) => {
          const item = document.createElement('div');
          item.className = 'file-item';
          item.innerHTML = `
            <div class="file-info">
              <div class="file-icon" style="color: #8b5cf6; background: rgba(139, 92, 246, 0.12); border-color: rgba(139, 92, 246, 0.3);">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
                  <circle cx="8.5" cy="8.5" r="1.5"></circle>
                  <polyline points="21 15 16 10 5 21"></polyline>
                </svg>
              </div>
              <div class="file-text">
                <div class="file-name" title="${file.name}">Hal ${index + 1}: ${file.name}</div>
                <div class="file-meta">${formatBytes(file.size)}</div>
              </div>
            </div>
            <div class="file-actions-group">
              <div class="order-btn-group">
                <button type="button" class="order-btn btn-up" title="Geser ke Atas" ${index === 0 ? 'disabled' : ''}>▲</button>
                <button type="button" class="order-btn btn-down" title="Geser ke Bawah" ${index === selectedImages.length - 1 ? 'disabled' : ''}>▼</button>
              </div>
              <button type="button" class="remove-file-btn" title="Hapus file ini">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line>
                </svg>
              </button>
            </div>
          `;

          item.querySelector('.btn-up').addEventListener('click', (e) => { e.stopPropagation(); moveItem(index, index - 1); });
          item.querySelector('.btn-down').addEventListener('click', (e) => { e.stopPropagation(); moveItem(index, index + 1); });
          item.querySelector('.remove-file-btn').addEventListener('click', (e) => {
            e.stopPropagation();
            selectedImages.splice(index, 1);
            updateUI();
          });

          filesList.appendChild(item);
        });
      } else {
        filesSection.style.display = 'none';
        convertBtn.disabled = true;
        convertBtnTxt.textContent = 'Konversi Gambar ke PDF';
      }
    }

    convertBtn.addEventListener('click', async () => {
      if (selectedImages.length === 0) return;

      resultCard.style.display = 'none';
      progressCont.style.display = 'flex';
      convertBtn.disabled = true;
      convertBtnTxt.textContent = 'Membuat Dokumen PDF...';
      progressFill.style.width = '45%';
      progressPct.textContent = '45%';
      progressText.textContent = 'Mengonversi dan menyusun resolusi lossless...';

      const fd = new FormData();
      selectedImages.forEach(f => fd.append('files', f));
      fd.append('page_size', pageSize);
      if (outputNameInp.value.trim()) {
        fd.append('output_name', outputNameInp.value.trim());
      }

      try {
        const res = await fetch('/api/image-to-pdf', { method: 'POST', body: fd });
        const data = await res.json();
        if (!res.ok || !data.success) {
          throw new Error(data.error || 'Gagal mengonversi gambar ke PDF');
        }

        progressFill.style.width = '100%';
        progressPct.textContent = '100%';
        progressText.textContent = 'PDF Berhasil Dibuat!';

        setTimeout(() => {
          progressCont.style.display = 'none';
          resultCard.style.display = 'block';
          resultSubtitle.textContent = `Berhasil mengemas ${data.total_images} gambar ke PDF (${data.size_kb} KB).`;
          downloadBtn.href = data.download_url;
          downloadBtn.setAttribute('download', data.filename);
          convertBtn.disabled = false;
          convertBtnTxt.textContent = `Konversi ${selectedImages.length} Gambar ke PDF`;
          resultCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }, 300);

      } catch (err) {
        progressCont.style.display = 'none';
        convertBtn.disabled = false;
        convertBtnTxt.textContent = `Konversi ${selectedImages.length} Gambar ke PDF`;
        alert('Terjadi kesalahan saat mengonversi gambar:\n' + err.message);
      }
    });

    if (folderBtn) {
      folderBtn.addEventListener('click', async () => {
        try { await fetch('/api/open-folder', { method: 'POST' }); } catch (e) { alert(e.message); }
      });
    }
  })();

}); // DOMContentLoaded
