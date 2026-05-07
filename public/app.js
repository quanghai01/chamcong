const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const browseBtn = document.getElementById('browseBtn');
const statusArea = document.getElementById('statusArea');
const spinner = document.getElementById('spinner');
const statusText = document.getElementById('statusText');
const downloadBtn = document.getElementById('downloadBtn');
const resetBtn = document.getElementById('resetBtn');

let currentBlobUrl = null;

// Handle Drag and Drop
uploadArea.addEventListener('dragover', (e) => {
  e.preventDefault();
  uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
  uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
  const files = e.dataTransfer.files;
  if (files.length) handleFile(files[0]);
});

// Handle Browse Button
browseBtn.addEventListener('click', (e) => {
  e.stopPropagation(); // prevent clicking uploadArea twice
  fileInput.click();
});

uploadArea.addEventListener('click', () => {
  fileInput.click();
});

fileInput.addEventListener('change', (e) => {
  if (e.target.files.length) handleFile(e.target.files[0]);
});

// Handle Reset
resetBtn.addEventListener('click', () => {
  uploadArea.classList.remove('hidden');
  statusArea.classList.add('hidden');
  fileInput.value = '';
  if (currentBlobUrl) URL.revokeObjectURL(currentBlobUrl);
});

async function handleFile(file) {
  if (!file.name.endsWith('.xls') && !file.name.endsWith('.xlsx')) {
    alert('Vui lòng chọn file Excel (.xls hoặc .xlsx)');
    return;
  }

  // Show processing UI
  uploadArea.classList.add('hidden');
  statusArea.classList.remove('hidden');
  spinner.classList.remove('hidden');
  downloadBtn.classList.add('hidden');
  resetBtn.classList.add('hidden');
  statusText.textContent = 'Đang xử lý dữ liệu chấm công...';
  statusText.style.color = 'var(--text-main)';

  const formData = new FormData();
  formData.append('file', file);

  try {
    const response = await fetch('/upload', {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      throw new Error(await response.text() || 'Lỗi server');
    }

    const blob = await response.blob();
    currentBlobUrl = URL.createObjectURL(blob);

    // Show success UI
    spinner.classList.add('hidden');
    statusText.textContent = '✅ Xử lý thành công!';
    statusText.style.color = 'var(--success)';
    
    downloadBtn.onclick = () => {
      const a = document.createElement('a');
      a.href = currentBlobUrl;
      const baseName = file.name.replace(/\.[^/.]+$/, "");
      a.download = `${baseName}_ketqua.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    };

    downloadBtn.classList.remove('hidden');
    resetBtn.classList.remove('hidden');

    // Auto trigger download
    downloadBtn.click();

  } catch (error) {
    spinner.classList.add('hidden');
    statusText.textContent = '❌ Lỗi: ' + error.message;
    statusText.style.color = '#ef4444';
    resetBtn.classList.remove('hidden');
  }
}
