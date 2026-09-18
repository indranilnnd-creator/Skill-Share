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
// wllama creates its worker as a module worker ({type:"module"}), which Chrome/Edge
// refuse to run from file:// pages (module scripts are CORS-restricted, and file://
// has a null origin). The worker code contains no import/export, so a classic worker
// works fine. Patch it here.
if (stripped.includes('{type:"module"}')) {
  var finalStripped = stripped.replace(/{type:"module"}/g, '{type:"classic"}');
  console.log('patched worker type module -> classic');
} else {
  console.error('WARNING: could not find {type:"module"} to patch');
  var finalStripped = stripped;
}

// On file:// pages, BOTH fetch() and atob() of large payloads inside a Worker hang.
// wllama's worker normally fetches the wasm via locateFile(), so loading hangs. Fix:
// decode the wasm bytes in the MAIN thread (where atob works) and pass the Uint8Array
// to the worker via the module.init message; the worker then sets Module.wasmBinary
// directly so the emscripten runtime never calls fetch().
const DECODE_WASM_INLINE =
  '(function(){var _b=atob(window.__WLLAMA_WASM_B64);var _u=new Uint8Array(_b.length);for(var _i=0;_i<_b.length;_i++)_u[_i]=_b.charCodeAt(_i);return _u;})()';

// Patch 1 (main thread): append the decoded wasm bytes as args[2] of module.init
const MODULE_INIT_ARGS = 'args:[new Blob([r],{type:"text/javascript"}),this.useAsyncFile]';
if (finalStripped.includes(MODULE_INIT_ARGS)) {
  finalStripped = finalStripped.replace(
    MODULE_INIT_ARGS,
    'args:[new Blob([r],{type:"text/javascript"}),this.useAsyncFile,' + DECODE_WASM_INLINE + ']'
  );
  console.log('patched moduleInit: decode wasm in main thread, pass as args[2]');
} else {
  console.error('WARNING: could not find module.init args to patch');
}

// Patch 2 (worker): set Module.wasmBinary from args[2] (the transferred bytes)
const GET_MODULE_CALL = 'Module = getWModuleConfig(argMainScriptBlob);';
if (finalStripped.includes(GET_MODULE_CALL)) {
  finalStripped = finalStripped.replace(
    GET_MODULE_CALL,
    'Module = getWModuleConfig(argMainScriptBlob); Module.wasmBinary = args[2];'
  );
  console.log('patched worker: set Module.wasmBinary from args[2]');
} else {
  console.error('WARNING: could not find "Module = getWModuleConfig(argMainScriptBlob);" to patch');
}

const wllamaBrowser = finalStripped + `
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
