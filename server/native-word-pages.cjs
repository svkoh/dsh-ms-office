/* global Buffer, process, __dirname, setTimeout, clearTimeout */
'use strict';
const path = require('node:path');
const crypto = require('node:crypto');
const { spawn } = require('node:child_process');
const RENDERER = 'Microsoft Word native Page.EnhMetaFileBits';
const MAX_BODY_CHARS = 1000000, MAX_OUTPUT_BYTES = 64 * 1024 * 1024, WATCHDOG_MS = 30000;
let rendering = false;
function canonicalBody(text) { return text.replaceAll('\r\u0007', '\t').replace(/\r\n?/g, '\n').replace(/\n+$/, ''); }
function documentFingerprint(text) { return crypto.createHash('sha256').update(canonicalBody(text), 'utf8').digest('hex'); }
function validateNativeWordPagesRequest(body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)
      || Object.keys(body).some(key => key !== 'bodyText' && key !== 'expectedPageCount')
      || typeof body.bodyText !== 'string' || body.bodyText.length > MAX_BODY_CHARS
      || !Number.isInteger(body.expectedPageCount) || body.expectedPageCount < 1 || body.expectedPageCount > 10) {
    return { ok: false, code: 'bad_args', error: 'Expected only bodyText (up to 1000000 characters) and expectedPageCount (integer 1..10)' };
  }
  return null;
}
function validateNativeResult(value, request) {
  if (!value || typeof value !== 'object' || value.ok !== true || !value.result) return false;
  const result = value.result;
  if (result.renderer !== RENDERER || result.documentFingerprint !== documentFingerprint(request.bodyText)
      || result.pageCount !== request.expectedPageCount || !Array.isArray(result.pages) || result.pages.length !== result.pageCount) return false;
  return result.pages.every((page, offset) => {
    if (!page || page.index !== offset + 1 || !Number.isInteger(page.width) || page.width < 1 || page.width > 2400
        || !Number.isInteger(page.height) || page.height < 1 || page.height > 10000
        || typeof page.base64 !== 'string' || !/^[A-Za-z0-9+/]+={0,2}$/.test(page.base64)) return false;
    const png = Buffer.from(page.base64, 'base64');
    return png.length >= 24 && png.subarray(0, 8).equals(Buffer.from('89504e470d0a1a0a', 'hex'))
      && png.toString('ascii', 12, 16) === 'IHDR' && png.readUInt32BE(16) === page.width && png.readUInt32BE(20) === page.height;
  });
}
// Request data travels only over stdin; all process paths and flags are plugin-owned.
// Options are test dependency injection and are never accepted from HTTP.
async function renderNativeWordPages(request, options = {}) {
  const invalid = validateNativeWordPagesRequest(request); if (invalid) return invalid;
  if ((options.platform || process.platform) !== 'win32') return { ok: false, code: 'requirement', error: 'Native Word page rendering requires Windows desktop Microsoft Word' };
  if (rendering) return { ok: false, code: 'busy', error: 'Another native Word page render is in progress' };
  rendering = true;
  try { return await new Promise(resolve => {
    let child, timer, settled = false, size = 0, stderrSize = 0; const chunks = [], errors = [];
    const finish = result => { if (settled) return; settled = true; if (timer) clearTimeout(timer); resolve(result); };
    const stopOwnChild = () => { try { if (child) child.kill(); } catch { /* Only our PowerShell child, never Word or a process tree. */ } };
    const executable = path.win32.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
    const helper = path.join(__dirname, 'native-word-pages.ps1');
    try { child = (options.spawn || spawn)(executable, ['-NoProfile', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', helper], { windowsHide: true, shell: false, stdio: ['pipe', 'pipe', 'pipe'] }); }
    catch { finish({ ok: false, code: 'requirement', error: 'Windows PowerShell native renderer could not start' }); return; }
    timer = setTimeout(() => { finish({ ok: false, code: 'timeout', error: 'Native Word page rendering exceeded the 30 second watchdog; only its PowerShell helper was stopped' }); stopOwnChild(); }, options.watchdogMs || WATCHDOG_MS);
    child.on('error', () => finish({ ok: false, code: 'requirement', error: 'Windows PowerShell native renderer could not start' }));
    child.stdout.on('data', chunk => {
      if (settled) return; size += chunk.length;
      if (size > MAX_OUTPUT_BYTES) { finish({ ok: false, code: 'execution', error: 'Native Word page output exceeded its bounded transport limit' }); stopOwnChild(); } else chunks.push(Buffer.from(chunk));
    });
    child.stderr.on('data', chunk => { if (stderrSize >= 16384) return; const part = Buffer.from(chunk).subarray(0, 16384 - stderrSize); errors.push(part); stderrSize += part.length; });
    child.on('close', exitCode => {
      if (settled) return; let output;
      try { output = JSON.parse(Buffer.concat(chunks).toString('utf8').replace(/^\uFEFF/, '')); }
      catch { finish({ ok: false, code: 'requirement', error: 'Native Word page API unavailable or helper failed', details: Buffer.concat(errors).toString('utf8') }); return; }
      if (output && output.ok === false && ['requirement', 'not_found', 'execution'].includes(output.code) && typeof output.error === 'string') finish({ ok: false, code: output.code, error: output.error });
      else if (exitCode === 0 && validateNativeResult(output, request)) finish(output);
      else finish({ ok: false, code: 'execution', error: 'Native Word page renderer returned an invalid or mismatched document result' });
    });
    child.stdin.on('error', () => { /* EPIPE may follow a native requirement error; its JSON/exit remains authoritative. */ });
    try { child.stdin.end(Buffer.from(JSON.stringify(request), 'utf8')); }
    catch { finish({ ok: false, code: 'execution', error: 'Native Word renderer input transport failed' }); stopOwnChild(); }
  }); } finally { rendering = false; }
}
module.exports = { renderNativeWordPages, validateNativeWordPagesRequest, canonicalBody, documentFingerprint, RENDERER };
