/* ==========================================================================
   HD Docx to PDF Converter - Client Logic
   ========================================================================== */

document.addEventListener('DOMContentLoaded', () => {
  // Elements
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');
  const browseBtn = document.getElementById('browse-btn');
  const selectedFilesSection = document.getElementById('selected-files-section');
  const filesList = document.getElementById('files-list');
  const fileCountSpan = document.getElementById('file-count');
  const clearAllBtn = document.getElementById('clear-all-btn');
  const convertBtn = document.getElementById('convert-btn');
  const convertBtnText = document.getElementById('convert-btn-text');

  const optOptimizePrint = document.getElementById('opt-optimize-print');
  const optPreserveFonts = document.getElementById('opt-preserve-fonts');
  const optBookmarks = document.getElementById('opt-bookmarks');

  const progressContainer = document.getElementById('progress-container');
  const progressBarFill = document.getElementById('progress-bar-fill');
  const progressStatusText = document.getElementById('progress-status-text');
  const progressPercentage = document.getElementById('progress-percentage');

  const resultsCard = document.getElementById('results-card');
  const resultsSubtitle = document.getElementById('results-subtitle');
  const convertedList = document.getElementById('converted-list');
  const downloadZipBtn = document.getElementById('download-zip-btn');
  const openFolderBtn = document.getElementById('open-folder-btn');

  // State
  let selectedFiles = [];

  // Drag and drop event listeners
  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('dragover');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('dragover');
    }, false);
  });

  dropZone.addEventListener('drop', (e) => {
    const droppedFiles = Array.from(e.dataTransfer.files);
    handleNewFiles(droppedFiles);
  });

  browseBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
  });

  dropZone.addEventListener('click', (e) => {
    if (e.target !== browseBtn) {
      fileInput.click();
    }
  });

  fileInput.addEventListener('change', (e) => {
    const pickedFiles = Array.from(e.target.files);
    handleNewFiles(pickedFiles);
    fileInput.value = ''; // Reset input agar file yang sama bisa dipilih ulang jika perlu
  });

  clearAllBtn.addEventListener('click', () => {
    selectedFiles = [];
    updateFileListUI();
  });

  function handleNewFiles(files) {
    const validExtensions = ['.docx', '.doc'];
    const addedFiles = [];

    files.forEach(file => {
      const ext = '.' + file.name.split('.').pop().toLowerCase();
      if (validExtensions.includes(ext)) {
        // Hindari duplikasi file dengan nama sama
        const alreadyExists = selectedFiles.some(f => f.name === file.name && f.size === file.size);
        if (!alreadyExists) {
          selectedFiles.push(file);
          addedFiles.push(file);
        }
      }
    });

    if (addedFiles.length === 0 && files.length > 0) {
      alert('Harap masukkan file berekstensi .docx atau .doc.');
    }

    updateFileListUI();
  }

  function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  function updateFileListUI() {
    filesList.innerHTML = '';
    fileCountSpan.textContent = selectedFiles.length;

    if (selectedFiles.length > 0) {
      selectedFilesSection.style.display = 'block';
      convertBtn.disabled = false;
      convertBtnText.textContent = `Mulai Konversi (${selectedFiles.length} File) ke HD PDF`;

      selectedFiles.forEach((file, index) => {
        const item = document.createElement('div');
        item.className = 'file-item';
        item.innerHTML = `
          <div class="file-info">
            <div class="file-icon">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
              </svg>
            </div>
            <div class="file-text">
              <div class="file-name" title="${file.name}">${file.name}</div>
              <div class="file-meta">${formatBytes(file.size)} &bull; Siap Konversi HD</div>
            </div>
          </div>
          <button type="button" class="remove-file-btn" data-index="${index}" title="Hapus file ini">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <line x1="18" y1="6" x2="6" y2="18"></line>
              <line x1="6" y1="6" x2="18" y2="18"></line>
            </svg>
          </button>
        `;

        item.querySelector('.remove-file-btn').addEventListener('click', (e) => {
          e.stopPropagation();
          selectedFiles.splice(index, 1);
          updateFileListUI();
        });

        filesList.appendChild(item);
      });
    } else {
      selectedFilesSection.style.display = 'none';
      convertBtn.disabled = true;
      convertBtnText.textContent = 'Mulai Konversi ke HD PDF';
    }
  }

  // Convert Button Action
  convertBtn.addEventListener('click', async () => {
    if (selectedFiles.length === 0) return;

    // Reset results & show progress
    resultsCard.style.display = 'none';
    progressContainer.style.display = 'flex';
    convertBtn.disabled = true;
    convertBtnText.textContent = 'Sedang Mengonversi ke HD...';

    progressBarFill.style.width = '15%';
    progressPercentage.textContent = '15%';
    progressStatusText.textContent = 'Menyiapkan engine konversi dan memuat file...';

    const formData = new FormData();
    selectedFiles.forEach(file => {
      formData.append('files', file);
    });

    formData.append('optimize_print', optOptimizePrint.checked ? 'true' : 'false');
    formData.append('preserve_fonts', optPreserveFonts.checked ? 'true' : 'false');
    formData.append('create_bookmarks', optBookmarks.checked ? 'true' : 'false');

    // Smooth progress ticker
    let progress = 15;
    const progressInterval = setInterval(() => {
      if (progress < 88) {
        progress += Math.floor(Math.random() * 8) + 2;
        if (progress > 88) progress = 88;
        progressBarFill.style.width = `${progress}%`;
        progressPercentage.textContent = `${progress}%`;

        if (progress > 40 && progress < 70) {
          progressStatusText.textContent = 'Word Engine sedang merender grafis & vektor pada resolusi 300+ DPI...';
        } else if (progress >= 70) {
          progressStatusText.textContent = 'Mengunci struktur PDF dan mempertahankan resolusi gambar asli...';
        }
      }
    }, 450);

    try {
      const response = await fetch('/api/convert', {
        method: 'POST',
        body: formData
      });

      clearInterval(progressInterval);
      const data = await response.json();

      if (!response.ok || !data.success) {
        throw new Error(data.error || (data.errors && data.errors[0]?.error) || 'Gagal melakukan konversi');
      }

      progressBarFill.style.width = '100%';
      progressPercentage.textContent = '100%';
      progressStatusText.textContent = 'Konversi selesai dengan sukses!';

      setTimeout(() => {
        progressContainer.style.display = 'none';
        renderResults(data);
        convertBtn.disabled = false;
        convertBtnText.textContent = `Mulai Konversi (${selectedFiles.length} File) ke HD PDF`;
      }, 500);

    } catch (err) {
      clearInterval(progressInterval);
      progressContainer.style.display = 'none';
      convertBtn.disabled = false;
      convertBtnText.textContent = `Mulai Konversi (${selectedFiles.length} File) ke HD PDF`;
      alert('Terjadi kesalahan saat konversi:\n' + err.message);
    }
  });

  function renderResults(data) {
    resultsCard.style.display = 'block';
    convertedList.innerHTML = '';

    const count = data.converted.length;
    resultsSubtitle.textContent = `Berhasil mengonversi ${count} dokumen menjadi PDF kualitas HD tanpa kompresi gambar.`;

    if (data.zip_url) {
      downloadZipBtn.href = data.zip_url;
      downloadZipBtn.style.display = 'inline-flex';
    } else {
      downloadZipBtn.style.display = 'none';
    }

    data.converted.forEach(item => {
      const div = document.createElement('div');
      div.className = 'converted-item';

      let imgMeta = '';
      if (item.image_info && item.image_info.image_count > 0) {
        imgMeta = `&bull; ${item.image_info.image_count} Gambar Terdeteksi (${item.image_info.estimated_quality})`;
      }

      div.innerHTML = `
        <div class="converted-info">
          <div class="pdf-icon">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
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
              <span>${imgMeta}</span>
            </div>
          </div>
        </div>
        <div>
          <a href="${item.download_url}" class="btn btn-primary" download>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            Unduh PDF
          </a>
        </div>
      `;
      convertedList.appendChild(div);
    });

    // Scroll to results
    resultsCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  // Open output folder in explorer
  openFolderBtn.addEventListener('click', async () => {
    try {
      const res = await fetch('/api/open-folder', { method: 'POST' });
      const d = await res.json();
      if (!d.success) {
        alert('Gagal membuka folder: ' + d.error);
      }
    } catch (e) {
      alert('Error saat membuka folder: ' + e.message);
    }
  });
});
