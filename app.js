// Use the globals provided by the libraries directly (do NOT redeclare them with
// const/let, since e.g. xlsx already declares `var XLSX` at top level, which would
// throw "Identifier 'XLSX' has already been declared").
const Docx = window.docx;
const jsPDF = (window.jspdf && window.jspdf.jsPDF) || (window.jsPDF);

const STORAGE_KEYS = {
  SKILLS: 'skillshare_skills',
  SETTINGS: 'skillshare_settings'
};

let wllama = null;
let modelLoaded = false;
let currentAbortController = null;
let files = [];
let skills = [];
let settings = {
  ctxSize: 4096,
  nThreads: 4,
  nGpuLayers: 99,
  temperature: 0.7,
  topP: 0.9,
  maxTokens: 2048
};
let lastStructuredResult = null;

// Configure pdf.js worker from the embedded base64 (avoids fetch() on file://)
function setupPdfWorker() {
  try {
    if (window.__PDF_WORKER_B64) {
      const bin = atob(window.__PDF_WORKER_B64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      const blob = new Blob([bytes], { type: 'text/javascript' });
      pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
    }
  } catch (e) {
    log('PDF worker setup failed, will use main-thread fallback: ' + e.message, 'warn');
  }
}

function log(message, type = 'info') {
  const logArea = document.getElementById('logArea');
  const line = document.createElement('div');
  line.className = `log-line log-${type}`;
  line.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
  logArea.appendChild(line);
  logArea.scrollTop = logArea.scrollHeight;
}

function setModelStatus(text, type = 'loading') {
  const el = document.getElementById('modelStatus');
  if (!el) return;
  el.className = `status-badge ${type}`;
  // Only update the text span, keeping the .dot indicator intact
  const textEl = document.getElementById('modelStatusText');
  if (textEl) {
    textEl.textContent = text;
  } else {
    el.textContent = text;
  }
}

function updateSkillCount() {
  document.getElementById('skillCount').textContent = skills.length;
}

function updateFileCount() {
  document.getElementById('fileCount').textContent = files.length;
}

function updateRunButton() {
  document.getElementById('runBtn').disabled = !(modelLoaded && files.length > 0);
}

function saveSkills() {
  localStorage.setItem(STORAGE_KEYS.SKILLS, JSON.stringify(skills));
}

function loadSkills() {
  try {
    const data = localStorage.getItem(STORAGE_KEYS.SKILLS);
    if (data) skills = JSON.parse(data);
  } catch (e) { log('Failed to load skills: ' + e.message, 'error'); }
  renderSkills();
}

function saveSettings() {
  localStorage.setItem(STORAGE_KEYS.SETTINGS, JSON.stringify(settings));
}

function loadSettings() {
  try {
    const data = localStorage.getItem(STORAGE_KEYS.SETTINGS);
    if (data) settings = { ...settings, ...JSON.parse(data) };
  } catch (e) { log('Failed to load settings: ' + e.message, 'error'); }
  applySettingsToUI();
}

function applySettingsToUI() {
  document.getElementById('ctxSize').value = settings.ctxSize;
  document.getElementById('nThreads').value = settings.nThreads;
  document.getElementById('nGpuLayers').value = settings.nGpuLayers;
  document.getElementById('temperature').value = settings.temperature;
  document.getElementById('topP').value = settings.topP;
  document.getElementById('maxTokens').value = settings.maxTokens;
  document.getElementById('tempValue').textContent = settings.temperature;
  document.getElementById('topPValue').textContent = settings.topP;
}

async function parseFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  const arrayBuffer = await file.arrayBuffer();
  let text = '';

  try {
    switch (ext) {
      case 'docx': {
        const result = await mammoth.extractRawText({ arrayBuffer });
        text = result.value;
        break;
      }
      case 'pdf': {
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const pages = [];
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const content = await page.getTextContent();
          pages.push(content.items.map(item => item.str).join(' '));
        }
        text = pages.join('\n\n');
        break;
      }
      case 'xlsx':
      case 'xls': {
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheets = workbook.SheetNames.map(name => {
          const sheet = workbook.Sheets[name];
          return XLSX.utils.sheet_to_csv(sheet);
        });
        text = sheets.join('\n\n---\n\n');
        break;
      }
      case 'csv': {
        const decoder = new TextDecoder('utf-8');
        text = decoder.decode(arrayBuffer);
        break;
      }
      case 'md':
      case 'txt': {
        const decoder = new TextDecoder('utf-8');
        text = decoder.decode(arrayBuffer);
        break;
      }
      default:
        throw new Error(`Unsupported file type: ${ext}`);
    }
    return { name: file.name, text, size: file.size, type: ext };
  } catch (e) {
    log(`Failed to parse ${file.name}: ${e.message}`, 'error');
    throw e;
  }
}

function chunkText(text, maxTokens = 2500) {
  const words = text.split(/\s+/);
  const chunks = [];
  let currentChunk = [];
  let currentTokens = 0;
  const tokensPerWord = 1.3;

  for (const word of words) {
    const wordTokens = Math.ceil(word.length / 4);
    if (currentTokens + wordTokens > maxTokens && currentChunk.length > 0) {
      chunks.push(currentChunk.join(' '));
      currentChunk = [word];
      currentTokens = wordTokens;
    } else {
      currentChunk.push(word);
      currentTokens += wordTokens;
    }
  }
  if (currentChunk.length > 0) {
    chunks.push(currentChunk.join(' '));
  }
  return chunks;
}

function renderSkills() {
  const list = document.getElementById('skillList');
  list.innerHTML = '';
  skills.forEach((skill, index) => {
    const item = document.createElement('div');
    item.className = 'skill-item';
    item.innerHTML = `
      <input type="checkbox" id="skill-${index}" ${skill.enabled ? 'checked' : ''}>
      <div class="skill-text">
        <strong>${skill.name}</strong>
        <div style="font-size:0.75rem; color:var(--text-muted); margin-top:2px;">${skill.prompt.substring(0, 100)}${skill.prompt.length > 100 ? '...' : ''}</div>
      </div>
      <div class="skill-actions">
        <button class="icon-btn" data-action="edit" title="Edit">✏️</button>
        <button class="icon-btn" data-action="delete" title="Delete">🗑️</button>
      </div>
    `;
    item.querySelector('input').addEventListener('change', (e) => {
      skills[index].enabled = e.target.checked;
      saveSkills();
    });
    item.querySelector('[data-action="edit"]').addEventListener('click', () => editSkill(index));
    item.querySelector('[data-action="delete"]').addEventListener('click', () => deleteSkill(index));
    list.appendChild(item);
  });
  updateSkillCount();
}

function editSkill(index) {
  const skill = skills[index];
  document.getElementById('skillName').value = skill.name;
  document.getElementById('skillPrompt').value = skill.prompt;
  document.getElementById('addSkillBtn').textContent = 'Update Skill';
  document.getElementById('addSkillBtn').dataset.editing = index;
  document.querySelector('[data-tab="skills"]').click();
}

function deleteSkill(index) {
  if (confirm(`Delete skill "${skills[index].name}"?`)) {
    skills.splice(index, 1);
    saveSkills();
    renderSkills();
  }
}

document.getElementById('addSkillBtn').addEventListener('click', () => {
  const name = document.getElementById('skillName').value.trim();
  const prompt = document.getElementById('skillPrompt').value.trim();
  if (!name || !prompt) {
    alert('Please enter both name and instructions');
    return;
  }
  const editingIndex = document.getElementById('addSkillBtn').dataset.editing;
  if (editingIndex !== undefined) {
    skills[editingIndex] = { name, prompt, enabled: skills[editingIndex].enabled };
    delete document.getElementById('addSkillBtn').dataset.editing;
    document.getElementById('addSkillBtn').textContent = 'Add Skill';
  } else {
    skills.push({ name, prompt, enabled: true });
  }
  saveSkills();
  renderSkills();
  document.getElementById('skillName').value = '';
  document.getElementById('skillPrompt').value = '';
});

document.getElementById('exportSkillsBtn').addEventListener('click', () => {
  const blob = new Blob([JSON.stringify(skills, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'skills.json';
  a.click();
  URL.revokeObjectURL(url);
});

document.getElementById('importSkillsBtn').addEventListener('click', () => {
  document.getElementById('importSkillsFile').click();
});

document.getElementById('importSkillsFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const text = await file.text();
    const imported = JSON.parse(text);
    if (Array.isArray(imported)) {
      skills = imported.map(s => ({ name: s.name, prompt: s.prompt, enabled: s.enabled !== false }));
      saveSkills();
      renderSkills();
      log(`Imported ${skills.length} skills`, 'success');
    }
  } catch (err) {
    log('Import failed: ' + err.message, 'error');
  }
  e.target.value = '';
});

function renderFiles() {
  const list = document.getElementById('fileList');
  list.innerHTML = '';
  files.forEach((file, index) => {
    const icon = getFileIcon(file.type);
    const item = document.createElement('div');
    item.className = 'file-item';
    item.innerHTML = `
      <span class="file-icon">${icon}</span>
      <div class="file-info">
        <div class="file-name">${file.name}</div>
        <div class="file-meta">${formatBytes(file.size)} • ${file.type.toUpperCase()} • ~${estimateTokens(file.text)} tokens</div>
      </div>
      <button class="file-remove" title="Remove">✕</button>
    `;
    item.querySelector('.file-remove').addEventListener('click', () => {
      files.splice(index, 1);
      renderFiles();
    });
    list.appendChild(item);
  });
  updateFileCount();
  updateRunButton();
}

function getFileIcon(type) {
  const icons = { docx: '📄', pdf: '📕', xlsx: '📊', xls: '📊', csv: '📋', md: '📝', txt: '📄' };
  return icons[type] || '📄';
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function estimateTokens(text) {
  return Math.ceil(text.length / 4);
}

const fileDrop = document.getElementById('fileDrop');
const fileInput = document.getElementById('fileInput');

fileDrop.addEventListener('click', () => fileInput.click());
fileDrop.addEventListener('dragover', (e) => { e.preventDefault(); fileDrop.classList.add('drag-over'); });
fileDrop.addEventListener('dragleave', () => fileDrop.classList.remove('drag-over'));
fileDrop.addEventListener('drop', async (e) => {
  e.preventDefault();
  fileDrop.classList.remove('drag-over');
  const droppedFiles = Array.from(e.dataTransfer.files);
  await handleFiles(droppedFiles);
});

fileInput.addEventListener('change', async (e) => {
  const selectedFiles = Array.from(e.target.files);
  await handleFiles(selectedFiles);
  e.target.value = '';
});

async function handleFiles(newFiles) {
  for (const file of newFiles) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['docx', 'pdf', 'md', 'txt', 'csv', 'xlsx', 'xls'].includes(ext)) {
      log(`Skipped ${file.name}: unsupported file type`, 'warn');
      continue;
    }
    if (files.some(f => f.name === file.name && f.size === file.size)) {
      log(`Skipped ${file.name}: already loaded`, 'warn');
      continue;
    }
    try {
      log(`Parsing ${file.name}...`, 'info');
      const parsed = await parseFile(file);
      files.push(parsed);
      log(`Loaded ${file.name} (${formatBytes(file.size)}, ~${estimateTokens(parsed.text)} tokens)`, 'success');
    } catch (err) {
      log(`Failed to load ${file.name}: ${err.message}`, 'error');
    }
  }
  renderFiles();
}

function getEnabledSkillsPrompt() {
  return skills.filter(s => s.enabled).map(s => `## Skill: ${s.name}\n${s.prompt}`).join('\n\n');
}

async function loadModel() {
  const modelFileInput = document.getElementById('modelFileInput');
  modelFileInput.click();
}

document.getElementById('loadModelBtn').addEventListener('click', loadModel);

document.getElementById('modelFileInput').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  await initializeWllama(file);
  e.target.value = '';
});

document.getElementById('reconnectModelBtn').addEventListener('click', async () => {
  const modelFileInput = document.getElementById('modelFileInput');
  modelFileInput.click();
});

document.getElementById('forgetModelBtn').addEventListener('click', async () => {
  if (confirm('Forget saved model? You will need to reload it next time.')) {
    await forgetModelHandle();
    setModelStatus('No model loaded', 'loading');
    modelLoaded = false;
    document.getElementById('runBtn').disabled = true;
    document.getElementById('reconnectModelBtn').disabled = true;
  }
});

async function getStoredModelHandle() {
  return null; // Handles can't be persisted reliably on file://; user will re-pick
}

async function storeModelHandle(fileHandle) {
  // No-op: handles can't be persisted on file://
  document.getElementById('reconnectModelBtn').disabled = false;
}

async function forgetModelHandle() {
  document.getElementById('reconnectModelBtn').disabled = true;
}

async function initializeWllama(modelFile) {
  try {
    log('Initializing wllama...', 'info');
    setModelStatus('Loading model...', 'loading');
    document.getElementById('loadModelBtn').disabled = true;

    // The wllama WASM binary is embedded and handled inside lib/wllama.browser.js
    // (decoded in the main thread and passed to the worker as Module.wasmBinary), so
    // the "default" path here is only a placeholder that is never actually fetched.
    // WebGPU is only used when available; otherwise fall back to CPU (n_gpu_layers = 0)
    const gpuLayers = (typeof navigator !== 'undefined' && navigator.gpu) ? settings.nGpuLayers : 0;

    wllama = new window.Wllama({
      default: 'wllama.wasm'
    });

    // loadModel accepts a File/Blob directly (avoids the .gguf URL requirement)
    await wllama.loadModel([modelFile], {
      n_ctx: settings.ctxSize,
      n_threads: settings.nThreads,
      n_gpu_layers: gpuLayers,
      progressCallback: ({ loaded, total }) => {
        const pct = total > 0 ? Math.round((loaded / total) * 100) : 0;
        setModelStatus(`Loading model... ${pct}%`, 'loading');
      }
    });

    modelLoaded = true;
    setModelStatus('Model ready', 'ready');
    updateRunButton();
    document.getElementById('reconnectModelBtn').disabled = false;
    log('Model loaded successfully', 'success');
  } catch (err) {
    log('Model load failed: ' + err.message, 'error');
    setModelStatus('Load failed: ' + err.message, 'error');
    document.getElementById('loadModelBtn').disabled = false;
  }
}

document.getElementById('runBtn').addEventListener('click', runProcessing);
document.getElementById('stopBtn').addEventListener('click', stopProcessing);
document.getElementById('clearOutputBtn').addEventListener('click', () => {
  document.getElementById('outputArea').textContent = 'Run a prompt to see output here...';
  document.getElementById('structuredPreview').textContent = 'No structured data yet. Run a prompt that produces JSON/CSV/table output.';
  lastStructuredResult = null;
});

async function runProcessing() {
  if (!modelLoaded) { alert('Please load a model first'); return; }
  if (files.length === 0) { alert('Please load at least one file'); return; }

  const prompt = document.getElementById('promptInput').value.trim();
  if (!prompt) { alert('Please enter a prompt'); return; }

  currentAbortController = new AbortController();
  document.getElementById('runBtn').disabled = true;
  document.getElementById('stopBtn').disabled = false;
  document.getElementById('progressContainer').style.display = 'block';
  document.getElementById('progressText').style.display = 'block';
  document.getElementById('outputArea').textContent = 'Processing...\n';

  const skillPrompt = getEnabledSkillsPrompt();
  const systemPrompt = skillPrompt
    ? `You are a helpful assistant. Follow these skill instructions:\n\n${skillPrompt}`
    : 'You are a helpful assistant.';

  try {
    const allChunks = [];
    for (const file of files) {
      const chunks = chunkText(file.text);
      allChunks.push({ file: file.name, chunks });
    }

    let combinedOutput = '';
    const totalChunks = allChunks.reduce((sum, f) => sum + f.chunks.length, 0);
    let processedChunks = 0;

    for (const { file, chunks } of allChunks) {
      for (let i = 0; i < chunks.length; i++) {
        if (currentAbortController.signal.aborted) throw new Error('Stopped by user');

        const userMessage = `File: ${file} (part ${i + 1} of ${chunks.length})\n\n${chunks[i]}\n\nTask: ${prompt}`;

        const progress = Math.round((processedChunks / totalChunks) * 100);
        document.getElementById('progressContainer').querySelector('.fill').style.width = progress + '%';
        document.getElementById('progressText').textContent = `Processing ${file} - part ${i + 1}/${chunks.length} (${progress}%)`;

        const response = await wllama.createChatCompletion({
          messages: [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: userMessage }
          ],
          max_tokens: settings.maxTokens,
          temperature: settings.temperature,
          top_p: settings.topP,
          stream: true,
          abortSignal: currentAbortController.signal
        });

        let partOutput = '';
        for await (const chunk of response) {
          if (currentAbortController.signal.aborted) throw new Error('Stopped by user');
          const token = extractToken(chunk);
          if (token) {
            partOutput += token;
            combinedOutput += token;
            document.getElementById('outputArea').textContent = combinedOutput;
            document.getElementById('outputArea').scrollTop = document.getElementById('outputArea').scrollHeight;
          }
        }
        processedChunks++;
      }
    }

    document.getElementById('progressContainer').querySelector('.fill').style.width = '100%';
    document.getElementById('progressText').textContent = 'Complete!';

    tryExtractStructured(combinedOutput);
    log('Processing complete', 'success');
  } catch (err) {
    if (err.message !== 'Stopped by user') {
      log('Processing error: ' + err.message, 'error');
      document.getElementById('outputArea').textContent += '\n\nError: ' + err.message;
    } else {
      log('Processing stopped by user', 'warn');
    }
  } finally {
    updateRunButton();
    document.getElementById('stopBtn').disabled = true;
    setTimeout(() => {
      document.getElementById('progressContainer').style.display = 'none';
      document.getElementById('progressText').style.display = 'none';
    }, 2000);
  }
}

function extractToken(chunk) {
  if (!chunk) return '';
  const choices = chunk.choices && chunk.choices[0];
  if (!choices) return '';
  const delta = choices.delta || {};
  if (typeof delta.content === 'string') return delta.content;
  if (typeof choices.message === 'object' && choices.message && typeof choices.message.content === 'string') return choices.message.content;
  if (typeof choices.text === 'string') return choices.text;
  if (typeof chunk.content === 'string') return chunk.content;
  return '';
}

function stopProcessing() {
  if (currentAbortController) currentAbortController.abort();
}

function tryExtractStructured(text) {
  const jsonMatch = text.match(/```(?:json)?\n([\s\S]*?)\n```/);
  if (jsonMatch) {
    try {
      const parsed = JSON.parse(jsonMatch[1]);
      lastStructuredResult = parsed;
      document.getElementById('structuredPreview').textContent = JSON.stringify(parsed, null, 2);
      log('Extracted structured JSON data', 'success');
      return;
    } catch (e) { /* not valid JSON */ }
  }

  const csvMatch = text.match(/```(?:csv)?\n([\s\S]*?)\n```/);
  if (csvMatch) {
    try {
      const parsed = XLSX.utils.sheet_to_json(XLSX.utils.aoa_to_sheet(csvMatch[1].trim().split('\n').map(r => r.split(',')).filter(r => r.length > 1)), { header: 1 });
      lastStructuredResult = parsed;
      document.getElementById('structuredPreview').textContent = JSON.stringify(parsed, null, 2);
      log('Extracted structured CSV data', 'success');
      return;
    } catch (e) { /* not valid CSV */ }
  }

  lastStructuredResult = null;
  document.getElementById('structuredPreview').textContent = 'No structured data detected. Output is plain text.';
}

document.getElementById('generateExportBtn').addEventListener('click', generateExport);
document.getElementById('previewExportBtn').addEventListener('click', previewExport);

async function generateExport() {
  const format = document.getElementById('exportFormat').value;
  const filename = document.getElementById('exportFilename').value || 'output';
  const source = document.getElementById('exportSource').value;
  
  let content = source === 'output' 
    ? document.getElementById('outputArea').textContent 
    : JSON.stringify(lastStructuredResult, null, 2);

  if (!content || content.includes('Run a prompt')) {
    alert('No content to export');
    return;
  }

  try {
    let blob, mimeType;
    
    switch (format) {
      case 'md':
        blob = new Blob([content], { type: 'text/markdown' });
        break;
      case 'txt':
        blob = new Blob([content], { type: 'text/plain' });
        break;
      case 'html':
        blob = new Blob([`<pre>${escapeHtml(content)}</pre>`], { type: 'text/html' });
        break;
      case 'csv':
        if (lastStructuredResult && Array.isArray(lastStructuredResult)) {
          const ws = XLSX.utils.json_to_sheet(lastStructuredResult);
          const csv = XLSX.utils.sheet_to_csv(ws);
          blob = new Blob([csv], { type: 'text/csv' });
        } else {
          blob = new Blob([content], { type: 'text/csv' });
        }
        break;
      case 'docx': {
        const doc = new Docx.Document({
          sections: [{ children: content.split('\n').map(line => new Docx.Paragraph(line)) }]
        });
        const buffer = await Docx.Packer.toBlob(doc);
        blob = buffer;
        break;
      }
      case 'xlsx': {
        const wb = XLSX.utils.book_new();
        if (lastStructuredResult && Array.isArray(lastStructuredResult)) {
          const ws = XLSX.utils.json_to_sheet(lastStructuredResult);
          XLSX.utils.book_append_sheet(wb, ws, 'Data');
        } else {
          const ws = XLSX.utils.aoa_to_sheet([['Content'], [content]]);
          XLSX.utils.book_append_sheet(wb, ws, 'Data');
        }
        const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
        blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        break;
      }
      case 'pdf': {
        const pdf = new jsPDF();
        const lines = pdf.splitTextToSize(content, 180);
        pdf.text(lines, 10, 10);
        blob = pdf.output('blob');
        break;
      }
    }
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${filename}.${format}`;
    a.click();
    URL.revokeObjectURL(url);
    log(`Exported ${filename}.${format}`, 'success');
  } catch (err) {
    log('Export failed: ' + err.message, 'error');
    alert('Export failed: ' + err.message);
  }
}

function previewExport() {
  const format = document.getElementById('exportFormat').value;
  const source = document.getElementById('exportSource').value;
  let content = source === 'output' 
    ? document.getElementById('outputArea').textContent 
    : JSON.stringify(lastStructuredResult, null, 2);
  
  if (!content || content.includes('Run a prompt')) {
    alert('No content to preview');
    return;
  }
  
  const preview = window.open('', '_blank');
  if (format === 'html') {
    preview.document.write(`<pre>${escapeHtml(content)}</pre>`);
  } else {
    preview.document.write(`<pre>${escapeHtml(content)}</pre>`);
  }
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

document.getElementById('saveSettingsBtn').addEventListener('click', () => {
  settings.ctxSize = parseInt(document.getElementById('ctxSize').value);
  settings.nThreads = parseInt(document.getElementById('nThreads').value);
  settings.nGpuLayers = parseInt(document.getElementById('nGpuLayers').value);
  settings.temperature = parseFloat(document.getElementById('temperature').value);
  settings.topP = parseFloat(document.getElementById('topP').value);
  settings.maxTokens = parseInt(document.getElementById('maxTokens').value);
  saveSettings();
  log('Settings saved', 'success');
});

document.getElementById('temperature').addEventListener('input', (e) => {
  document.getElementById('tempValue').textContent = e.target.value;
});
document.getElementById('topP').addEventListener('input', (e) => {
  document.getElementById('topPValue').textContent = e.target.value;
});

document.getElementById('clearAllDataBtn').addEventListener('click', () => {
  if (confirm('Clear ALL local data? This includes skills, settings, and model handle.')) {
    localStorage.clear();
    skills = [];
    files = [];
    settings = { ctxSize: 4096, nThreads: 4, nGpuLayers: 99, temperature: 0.7, topP: 0.9, maxTokens: 2048 };
    renderSkills();
    renderFiles();
    applySettingsToUI();
    setModelStatus('No model loaded', 'loading');
    modelLoaded = false;
    document.getElementById('runBtn').disabled = true;
    document.getElementById('reconnectModelBtn').disabled = true;
    log('All local data cleared', 'warn');
  }
});

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => {
      b.classList.remove('active');
      b.setAttribute('aria-selected', 'false');
    });
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    
    btn.classList.add('active');
    btn.setAttribute('aria-selected', 'true');
    document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
  });
});

async function init() {
  setupPdfWorker();
  loadSkills();
  loadSettings();
  log('Skill Share initialized. Load a model to begin.', 'info');
}

init();