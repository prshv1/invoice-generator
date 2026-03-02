/* ============================================
   ChallanAI — App Logic
   ============================================ */

const BASE_URL = '/api';

// DOM refs
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileNameEl = document.getElementById('fileName');
const invNum = document.getElementById('invNum');
const outputFormat = document.getElementById('outputFormat');
const pdfToggle = document.getElementById('pdfToggle');
const generateBtn = document.getElementById('generateBtn');
const downloadBtn = document.getElementById('downloadBtn');
const statusEl = document.getElementById('status');

let selectedFile = null;
let downloadBlob = null;
let downloadName = '';

// ── Toggle ↔ Select sync ──
pdfToggle.addEventListener('change', () => {
  outputFormat.value = pdfToggle.checked ? 'pdf' : 'xlsx';
});

outputFormat.addEventListener('change', () => {
  pdfToggle.checked = outputFormat.value === 'pdf';
});

// ── Drag & Drop ──
dropzone.addEventListener('click', () => fileInput.click());

dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropzone.classList.add('dragover');
});

dropzone.addEventListener('dragleave', () => {
  dropzone.classList.remove('dragover');
});

dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files.length) handleFile(fileInput.files[0]);
});

// ── File handling ──
function handleFile(file) {
  selectedFile = file;
  fileNameEl.textContent = file.name;
  fileNameEl.style.display = 'flex';
  dropzone.classList.add('has-file');
  generateBtn.disabled = false;
  downloadBtn.style.display = 'none';
  setStatus('', '');
}

// ── Status display ──
function setStatus(msg, type) {
  statusEl.className = 'status' + (type ? ' ' + type : '');
  statusEl.innerHTML = msg;
}

// ── Endpoint routing ──
function getEndpoint(file, format) {
  const ext = file.name.split('.').pop().toLowerCase();
  const isImage = ['jpg', 'jpeg', 'png', 'bmp', 'tiff', 'tif', 'webp'].includes(ext);
  const isZip = ext === 'zip';

  if (isZip) return '/batch';
  if (isImage && format === 'pdf') return '/generate-from-image-pdf';
  if (isImage) return '/generate-from-image';
  if (format === 'pdf') return '/generate-pdf';
  return '/generate';
}

// ── Generate invoice ──
generateBtn.addEventListener('click', async () => {
  if (!selectedFile) return;

  const format = outputFormat.value;
  const endpoint = getEndpoint(selectedFile, format);
  const num = parseInt(invNum.value) || 1;

  const formData = new FormData();
  formData.append('file', selectedFile);

  if (endpoint === '/batch') {
    formData.append('start_num', num);
  } else {
    formData.append('inv_num', num);
  }

  generateBtn.disabled = true;
  downloadBtn.style.display = 'none';
  setStatus('<span class="spinner"></span>Processing your file…', 'loading');

  try {
    const resp = await fetch(BASE_URL + endpoint, {
      method: 'POST',
      body: formData,
    });

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({ detail: 'Server error' }));
      throw new Error(err.detail || `HTTP ${resp.status}`);
    }

    downloadBlob = await resp.blob();
    const contentDisp = resp.headers.get('Content-Disposition') || '';
    const match = contentDisp.match(/filename=(.+)/);
    downloadName = match
      ? match[1]
      : `Invoice_${num}.${endpoint === '/batch' ? 'zip' : format}`;

    setStatus('✓ Invoice generated successfully', 'success');
    downloadBtn.style.display = 'block';
  } catch (err) {
    setStatus('✕ ' + err.message, 'error');
  } finally {
    generateBtn.disabled = false;
  }
});

// ── Download ──
downloadBtn.addEventListener('click', () => {
  if (!downloadBlob) return;
  const url = URL.createObjectURL(downloadBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = downloadName;
  a.click();
  URL.revokeObjectURL(url);
});
