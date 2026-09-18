const fs = require('fs');
const path = require('path');

const libDir = path.join(__dirname, 'lib');

// 1. Build wllama.browser.js from wllama.min.js (strip ESM exports, expose globals)
const wllamaMin = fs.readFileSync(path.join(libDir, 'wllama.min.js'), 'utf8');
// Remove the trailing export statement
const exportRegex = /export\{[^}]*\};?\s*$/;
if (!exportRegex.test(wllamaMin)) {
  console.error('ERROR: could not find trailing export statement in wllama.min.js');
  process.exit(1);
}
const stripped = wllamaMin.replace(exportRegex, '');
const wllamaBrowser = stripped + `
window.Wllama = Wllama;
window.CacheManager = CacheManager;
window.ModelManager = ModelManager;
window.Model = Model;
window.WllamaError = WllamaError;
window.WllamaAbortError = WllamaAbortError;
window.WllamaRuntimeError = WllamaRuntimeError;
window.LogLevel = LogLevel;
window.LoggerWithoutDebug = LoggerWithoutDebug;
`;
fs.writeFileSync(path.join(libDir, 'wllama.browser.js'), wllamaBrowser);
console.log('wllama.browser.js written (' + (wllamaBrowser.length/1024).toFixed(0) + ' KB)');

// 2. Base64-encode wllama.wasm into wllama.wasm.js
const wasmBuf = fs.readFileSync(path.join(libDir, 'wllama.wasm'));
const wasmB64 = wasmBuf.toString('base64');
fs.writeFileSync(
  path.join(libDir, 'wllama.wasm.js'),
  'window.__WLLAMA_WASM_B64="' + wasmB64 + '";'
);
console.log('wllama.wasm.js written (' + (wasmB64.length/1024).toFixed(0) + ' KB base64)');

// 3. Base64-encode pdf.worker.min.js into pdf.worker.js
const workerBuf = fs.readFileSync(path.join(libDir, 'pdf.worker.min.js'));
const workerB64 = workerBuf.toString('base64');
fs.writeFileSync(
  path.join(libDir, 'pdf.worker.js'),
  'window.__PDF_WORKER_B64="' + workerB64 + '";'
);
console.log('pdf.worker.js written (' + (workerB64.length/1024).toFixed(0) + ' KB base64)');

console.log('Build complete.');
