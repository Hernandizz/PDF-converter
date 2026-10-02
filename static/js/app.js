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

}); // DOMContentLoaded
