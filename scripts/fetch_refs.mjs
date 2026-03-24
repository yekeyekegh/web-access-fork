#!/usr/bin/env node
/**
 * fetch_refs.mjs — 通用 Reference PDF 下载器
 *
 * 用法：
 *   node fetch_refs.mjs <excel_path> <start> <end> [output_dir]
 *
 * 示例：
 *   node fetch_refs.mjs "D:/path/to/reference_list.xlsx" 1 10
 *   node fetch_refs.mjs "D:/path/to/reference_list.xlsx" 21 30 "D:/output"
 *
 * Excel 格式预期：
 *   第一行为标题行（"Works cited" 等），从第二行开始是条目
 *   每条格式：N. Title, accessed Date, https://url
 *
 * 支持的 URL 类型（自动识别）：
 *   direct_pdf   — URL 以 .pdf 结尾或路径含 /pdf/，直接下载原始文件
 *   researchgate — researchgate.net，CDP 打开后等待 Cloudflare 通过
 *   youtube      — YouTube 视频，用 yt-dlp 下载字幕并转为 PDF
 *   html         — 普通网页，CDP printToPDF
 *
 *   注：下载后若发现是 .doc/.docx 文件，自动用 Word COM 转为 PDF（Windows）
 *
 * 依赖：
 *   - Node.js 22+（原生 WebSocket）
 *   - Python 3 + openpyxl（Excel 读取）
 *   - Chrome 开启远程调试（端口 9222）
 *   - yt-dlp（YouTube 字幕，pip install yt-dlp）
 *   - win32com / pywin32（Word 转 PDF，pip install pywin32）
 */

import fs from 'fs';
import path from 'path';
import https from 'https';
import http from 'http';
import { spawnSync } from 'child_process';

// ═══════════════════════════════════════════════════════════
// CLI 参数解析
// ═══════════════════════════════════════════════════════════
const args = process.argv.slice(2);
if (args.length < 3) {
  console.log('用法: node fetch_refs.mjs <excel_path> <start> <end> [output_dir]');
  console.log('示例: node fetch_refs.mjs "reference_list.xlsx" 1 10');
  process.exit(0);
}
const EXCEL_PATH = path.resolve(args[0]);
const START_NUM = parseInt(args[1]);
const END_NUM = parseInt(args[2]);
const OUT_DIR = args[3] ? path.resolve(args[3]) : path.dirname(EXCEL_PATH);

if (!fs.existsSync(EXCEL_PATH)) {
  console.error('Excel 文件不存在:', EXCEL_PATH);
  process.exit(1);
}
fs.mkdirSync(OUT_DIR, { recursive: true });

console.log(`读取: ${EXCEL_PATH}`);
console.log(`范围: 第 ${START_NUM} ~ ${END_NUM} 条`);
console.log(`输出: ${OUT_DIR}\n`);

// ═══════════════════════════════════════════════════════════
// Excel 读取（调用 Python）
// ═══════════════════════════════════════════════════════════
function readExcelRows(excelPath, start, end) {
  const minRow = start + 1;  // Excel 第一行是标题行
  const maxRow = end + 1;
  const uid = `${process.pid}_${Date.now()}`;
  const tmpPy  = path.join(process.env.TEMP || 'C:/Windows/Temp', `read_excel_${uid}.py`);
  const tmpOut = path.join(process.env.TEMP || 'C:/Windows/Temp', `excel_rows_${uid}.json`);
  const pyScript = `# -*- coding: utf-8 -*-
import openpyxl, json, sys
wb = openpyxl.load_workbook(sys.argv[1])
ws = wb.active
rows = []
for row in ws.iter_rows(min_row=${minRow}, max_row=${maxRow}, values_only=True):
    if row[0]:
        rows.append(str(row[0]))
with open(sys.argv[2], 'w', encoding='utf-8') as f:
    json.dump(rows, f, ensure_ascii=False)
`;
  try {
    fs.writeFileSync(tmpPy, pyScript, 'utf-8');
    const r = spawnSync('python', [tmpPy, excelPath, tmpOut], { timeout: 10000, encoding: 'utf-8' });
    if (r.status !== 0) throw new Error(r.stderr || r.stdout || 'python exited with ' + r.status);
    const result = JSON.parse(fs.readFileSync(tmpOut, 'utf-8'));
    fs.unlinkSync(tmpPy);
    fs.unlinkSync(tmpOut);
    return result;
  } catch (e) {
    console.error('读取 Excel 失败:', e.message);
    process.exit(1);
  }
}

// ═══════════════════════════════════════════════════════════
// 条目解析
// ═══════════════════════════════════════════════════════════
function parseEntry(raw) {
  // 格式: "N. Title, accessed Date, https://url"
  // URL 提取：取最后一个 http(s):// 开头的段
  const urlMatch = raw.match(/(https?:\/\/[^\s,]+)(?:\s*$)/);
  const url = urlMatch ? urlMatch[1].trim() : null;

  // 去掉 URL 和 ", accessed ..." 部分
  let titlePart = raw;
  if (url) titlePart = titlePart.substring(0, titlePart.lastIndexOf(url)).replace(/,?\s*$/, '').trim();
  titlePart = titlePart.replace(/,\s*accessed\s+.+$/, '').trim();

  // 提取序号和标题
  const numMatch = titlePart.match(/^(\d+)\.\s*/);
  const num = numMatch ? parseInt(numMatch[1]) : 0;
  const title = numMatch ? titlePart.substring(numMatch[0].length).trim() : titlePart;

  return { num, title, url };
}

function sanitizeFilename(name) {
  return name
    .replace(/[\\/:*?"<>|']/g, '')   // 删除 Windows 非法字符和单引号
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 200);              // 限制长度
}

// ═══════════════════════════════════════════════════════════
// URL 类型自动识别
// ═══════════════════════════════════════════════════════════
function detectType(url) {
  if (!url) return 'html';
  const u = url.toLowerCase();
  if (u.includes('researchgate.net')) return 'researchgate';
  if (u.includes('pmc.ncbi.nlm.nih.gov')) return 'pmc';
  if (u.includes('frontiersin.org')) return 'frontiersin';
  if (u.includes('pubmed.ncbi.nlm.nih.gov')) return 'pubmed';
  if (u.includes('youtube.com') || u.includes('youtu.be')) return 'youtube';
  if (
    u.endsWith('.pdf') ||
    u.includes('/pdf/') ||
    u.includes('/pdfs/') ||
    u.includes('fulltext/ej') ||     // ERIC
    u.includes('/view/delivery/') || // ExLibris Alma 大学文件库（可能重定向到 .doc/.docx/.pdf）
    (u.includes('wp-content/uploads') && u.includes('.pdf'))
  ) return 'direct_pdf';
  return 'html';
}

// ═══════════════════════════════════════════════════════════
// Word (.doc/.docx) → PDF 转换（调用本机 Word COM，Windows）
// ═══════════════════════════════════════════════════════════
function convertWordToPdf(docPath, pdfPath) {
  const tmpPy = path.join(process.env.TEMP || 'C:/Windows/Temp', `word2pdf_${process.pid}.py`);
  const script = `# -*- coding: utf-8 -*-
import win32com.client, os, sys
doc_path = sys.argv[1]
pdf_path = sys.argv[2]
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
try:
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
finally:
    word.Quit()
`;
  fs.writeFileSync(tmpPy, script, 'utf-8');
  const r = spawnSync('python', [tmpPy, docPath, pdfPath], { timeout: 60000, encoding: 'utf-8' });
  try { fs.unlinkSync(tmpPy); } catch {}
  if (r.status !== 0) throw new Error('Word COM 转换失败: ' + (r.stderr || r.stdout));
}

// ═══════════════════════════════════════════════════════════
// Chrome CDP WebSocket URL（动态读取）
// ═══════════════════════════════════════════════════════════
function getChromeWsUrl() {
  const localAppData = process.env.LOCALAPPDATA || `C:/Users/${process.env.USERNAME}/AppData/Local`;
  const candidates = [
    `${localAppData}/Google/Chrome/User Data/DevToolsActivePort`,
    `${localAppData}/Google/Chrome Beta/User Data/DevToolsActivePort`,
    `${localAppData}/Chromium/User Data/DevToolsActivePort`,
  ];
  for (const p of candidates) {
    try {
      const lines = fs.readFileSync(p, 'utf-8').trim().split('\n');
      const port = parseInt(lines[0]);
      const wsPath = lines[1]?.trim() || '/devtools/browser';
      if (port > 0) return `ws://127.0.0.1:${port}${wsPath}`;
    } catch {}
  }
  return 'ws://127.0.0.1:9222/devtools/browser';
}

// ═══════════════════════════════════════════════════════════
// 直接 PDF 下载（不经 Chrome PDF Viewer）
// ═══════════════════════════════════════════════════════════
function downloadPdf(url, dest, redirects = 0) {
  return new Promise((resolve, reject) => {
    if (redirects > 8) return reject(new Error('Too many redirects'));
    const mod = url.startsWith('https') ? https : http;
    const req = mod.get(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
      },
      timeout: 30000,
    }, (res) => {
      // 跟随重定向（处理相对 URL）
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        res.resume();
        const redirectUrl = new URL(res.headers.location, url).href;
        return downloadPdf(redirectUrl, dest, redirects + 1).then(resolve).catch(reject);
      }
      if (res.statusCode !== 200) {
        res.resume();
        return reject(new Error(`HTTP ${res.statusCode}`));
      }
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        const buf = Buffer.concat(chunks);
        fs.writeFileSync(dest, buf);
        resolve(buf.length);
      });
      res.on('error', reject);
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('Timeout')); });
  });
}

// ═══════════════════════════════════════════════════════════
// CDP 工具
// ═══════════════════════════════════════════════════════════
let ws, msgId = 1;
const pending = new Map();
const evtListeners = new Map();

function connectCDP() {
  const wsUrl = getChromeWsUrl();
  console.log('[CDP] 连接:', wsUrl);
  return new Promise((resolve, reject) => {
    ws = new WebSocket(wsUrl);
    ws.onopen = () => { console.log('[CDP] 已连接\n'); resolve(); };
    ws.onerror = e => reject(new Error('WebSocket: ' + (e.message || e)));
    ws.onmessage = evt => {
      const msg = JSON.parse(evt.data);
      if (msg.id && pending.has(msg.id)) {
        const { resolve, reject } = pending.get(msg.id);
        pending.delete(msg.id);
        if (msg.error) reject(new Error(JSON.stringify(msg.error)));
        else resolve(msg.result);
      } else if (msg.method) {
        const key = (msg.sessionId || '') + ':' + msg.method;
        for (const fn of (evtListeners.get(key) || [])) fn(msg.params);
      }
    };
  });
}

function cdp(method, params = {}, sessionId = null, timeout = 30000) {
  return new Promise((resolve, reject) => {
    const id = msgId++;
    const msg = { id, method, params };
    if (sessionId) msg.sessionId = sessionId;
    const timer = setTimeout(() => {
      pending.delete(id);
      reject(new Error(`CDP timeout: ${method} (id=${id})`));
    }, timeout);
    pending.set(id, {
      resolve: v => { clearTimeout(timer); resolve(v); },
      reject: e => { clearTimeout(timer); reject(e); },
    });
    ws.send(JSON.stringify(msg));
  });
}

function once(method, sessionId = '') {
  return new Promise(resolve => {
    const key = sessionId + ':' + method;
    const arr = evtListeners.get(key) || [];
    const handler = params => {
      evtListeners.set(key, (evtListeners.get(key) || []).filter(f => f !== handler));
      resolve(params);
    };
    arr.push(handler);
    evtListeners.set(key, arr);
  });
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function openTab(url) {
  const { targetId } = await cdp('Target.createTarget', { url: 'about:blank', background: true });
  const { sessionId } = await cdp('Target.attachToTarget', { targetId, flatten: true });
  await cdp('Page.enable', {}, sessionId);
  return { targetId, sessionId };
}

async function navigateAndWait(sessionId, url, timeout = 20000) {
  const loadP = once('Page.loadEventFired', sessionId);
  await cdp('Page.navigate', { url }, sessionId);
  await Promise.race([loadP, sleep(timeout)]);
}

// 等待页面加载 + 网络空闲（适合 JS-heavy SPA 页面）
// 先等 loadEventFired，再等 networkIdle（无未决请求持续 500ms），最多等 idleTimeout
async function navigateAndWaitIdle(sessionId, url, loadTimeout = 20000, idleTimeout = 15000) {
  await cdp('Page.setLifecycleEventsEnabled', { enabled: true }, sessionId);
  const loadP = once('Page.loadEventFired', sessionId);
  const idleP = new Promise(resolve => {
    const key = sessionId + ':Page.lifecycleEvent';
    const arr = evtListeners.get(key) || [];
    const handler = params => {
      if (params?.name === 'networkIdle') {
        evtListeners.set(key, (evtListeners.get(key) || []).filter(f => f !== handler));
        resolve();
      }
    };
    arr.push(handler);
    evtListeners.set(key, arr);
  });
  await cdp('Page.navigate', { url }, sessionId);
  await Promise.race([loadP, sleep(loadTimeout)]);
  await Promise.race([idleP, sleep(idleTimeout)]);
}

async function printToPdf(sessionId, outPath) {
  const r = await cdp('Page.printToPDF', {
    printBackground: true,
    preferCSSPageSize: true,
    marginTop: 0.4, marginBottom: 0.4,
    marginLeft: 0.4, marginRight: 0.4,
  }, sessionId);
  const buf = Buffer.from(r.data, 'base64');
  fs.writeFileSync(outPath, buf);
  return buf.length;
}

// ═══════════════════════════════════════════════════════════
// Readability 文章提取 → 清洁 HTML → PDF
// 适用场景：含侧边栏/多栏布局的新闻、博客、期刊 landing page。
// 检测逻辑：Readability 能解析出 >2000 字符内容则视为"文章页"并启用。
// ═══════════════════════════════════════════════════════════
async function printWithReadability(sessionId, outPath) {
  // 注入 Mozilla Readability，解析文章主体
  const expr = `(async () => {
    if (typeof Readability === 'undefined') {
      await new Promise((res, rej) => {
        const s = document.createElement('script');
        s.src = 'https://cdn.jsdelivr.net/npm/@mozilla/readability/Readability.js';
        s.onload = res; s.onerror = rej;
        document.head.appendChild(s);
      });
    }
    const reader = new Readability(document.cloneNode(true));
    const a = reader.parse();
    if (!a || !a.content || a.content.length < 2000) return null;
    return JSON.stringify({title: a.title, byline: a.byline, content: a.content, siteName: a.siteName});
  })()`;
  const res = await cdp('Runtime.evaluate', {
    expression: expr, awaitPromise: true, returnByValue: true, timeout: 30000,
  }, sessionId);
  const article = JSON.parse(res?.result?.value || 'null');
  if (!article) return false;  // Readability 不适用，调用方回退到 printToPdf

  const e = s => (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const css = [
    'body{font-family:Georgia,serif;font-size:14pt;line-height:1.65;max-width:680px;margin:0 auto;padding:20px 40px;color:#111}',
    'h1{font-size:20pt;line-height:1.3;margin-bottom:6px}',
    '.byline{color:#555;font-size:11pt;margin-bottom:18px;border-bottom:1px solid #ddd;padding-bottom:12px}',
    'img{max-width:100%;height:auto;display:block;margin:14px auto}',
    'figure{margin:14px 0}figcaption{font-size:10pt;color:#666;text-align:center}',
    'blockquote{border-left:3px solid #ccc;margin-left:0;padding-left:14px;color:#444;font-style:italic}',
    'a{color:#333;text-decoration:none}',
    'table{border-collapse:collapse;width:100%;font-size:11pt}td,th{border:1px solid #ddd;padding:4px 8px}',
  ].join('');
  const bylineHtml = article.byline
    ? '<p class="byline">' + e(article.byline) + (article.siteName ? ' \xb7 ' + e(article.siteName) : '') + '</p>'
    : '';
  const html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + e(article.title) + '</title>'
    + '<style>' + css + '</style></head><body>'
    + '<h1>' + e(article.title) + '</h1>' + bylineHtml
    + article.content + '</body></html>';

  const uid = process.pid + '_' + Date.now();
  const tmpHtml = (process.env.TEMP || 'C:/Windows/Temp') + '\\readability_' + uid + '.html';
  fs.writeFileSync(tmpHtml, html, 'utf-8');
  const fileUrl = 'file:///' + tmpHtml.split('\\').join('/');

  const loadP = once('Page.loadEventFired', sessionId);
  await cdp('Page.navigate', { url: fileUrl }, sessionId);
  await Promise.race([loadP, sleep(5000)]);
  await sleep(300);

  const r = await cdp('Page.printToPDF', {
    printBackground: true,
    preferCSSPageSize: false,
    paperWidth: 8.27, paperHeight: 11.69,  // A4
    marginTop: 0.6, marginBottom: 0.6, marginLeft: 0.6, marginRight: 0.6,
  }, sessionId);
  const buf = Buffer.from(r.data, 'base64');
  fs.writeFileSync(outPath, buf);
  try { fs.unlinkSync(tmpHtml); } catch {}
  return buf.length;
}

// ═══════════════════════════════════════════════════════════
// Network.loadNetworkResource 下载 PDF
// 用 CDP 直接加载资源（携带浏览器 cookies），绕过 JS CORS 限制
// ═══════════════════════════════════════════════════════════
async function downloadPdfViaNetworkResource(sessionId, pdfUrl, outPath) {
  const { frameTree } = await cdp('Page.getFrameTree', {}, sessionId);
  const frameId = frameTree.frame.id;
  const { resource } = await cdp('Network.loadNetworkResource', {
    frameId,
    url: pdfUrl,
    options: { disableCache: false, includeCredentials: true },
  }, sessionId);
  if (!resource.success) throw new Error(`loadNetworkResource 失败 HTTP ${resource.httpStatusCode}`);
  const chunks = [];
  while (true) {
    const { data, eof } = await cdp('IO.read', { handle: resource.stream, size: 131072 }, sessionId);
    chunks.push(Buffer.from(data, 'base64'));
    if (eof) break;
  }
  await cdp('IO.close', { handle: resource.stream }, sessionId).catch(() => {});
  const buf = Buffer.concat(chunks);
  fs.writeFileSync(outPath, buf);
  return buf.length;
}

// ═══════════════════════════════════════════════════════════
// In-page fetch 下载（利用浏览器 session/cookies，绕过 CF TLS 指纹检测）
// ═══════════════════════════════════════════════════════════
async function downloadViaInPageFetch(sessionId, pdfUrl, outPath) {
  // 在浏览器上下文内 fetch，使用真实 Chrome 的 cookies 和 TLS 指纹
  const sizeRes = await cdp('Runtime.evaluate', {
    expression: `(async () => {
      window.__dl_buf = null;
      const r = await fetch(${JSON.stringify(pdfUrl)}, {credentials: 'include'});
      if (!r.ok) return JSON.stringify({error: r.status});
      const ab = await r.arrayBuffer();
      window.__dl_buf = ab;
      return JSON.stringify({size: new Uint8Array(ab).byteLength});
    })()`,
    awaitPromise: true,
    returnByValue: true,
    timeout: 120000,
  }, sessionId);
  const info = JSON.parse(sizeRes?.result?.value || '{}');
  if (info.error) throw new Error('In-page fetch HTTP ' + info.error);
  const totalSize = info.size;
  console.log(`  PDF size: ${(totalSize / 1024).toFixed(0)} KB`);

  // 分 512KB 块读回，避免 base64 字符串超限
  const CHUNK = 512 * 1024;
  const chunks = [];
  for (let offset = 0; offset < totalSize; offset += CHUNK) {
    const end = Math.min(offset + CHUNK, totalSize);
    const b64Res = await cdp('Runtime.evaluate', {
      expression: `(() => {
        const arr = new Uint8Array(window.__dl_buf, ${offset}, ${end - offset});
        let s = '';
        for (let i = 0; i < arr.length; i++) s += String.fromCharCode(arr[i]);
        return btoa(s);
      })()`,
      returnByValue: true,
    }, sessionId);
    chunks.push(Buffer.from(b64Res?.result?.value, 'base64'));
  }
  const buf = Buffer.concat(chunks);
  fs.writeFileSync(outPath, buf);
  await cdp('Runtime.evaluate', { expression: 'window.__dl_buf = null', returnByValue: true }, sessionId);
  return buf.length;
}

// Cloudflare 拦截检测（用于 ResearchGate 等）
async function waitForCloudflare(sessionId, maxWaitMs = 35000) {
  const start = Date.now();
  while (Date.now() - start < maxWaitMs) {
    await sleep(3000);
    const r = await cdp('Runtime.evaluate', {
      expression: `JSON.stringify({
        cf: document.title === 'Just a moment...' ||
            !!document.querySelector('#challenge-form'),
        title: document.title.substring(0, 80),
        hasContent: !!(document.querySelector('main, article, h1, .publication-detail, .research-detail')),
      })`,
      returnByValue: true,
    }, sessionId);
    const s = JSON.parse(r.result?.value || '{}');
    const elapsed = ((Date.now() - start) / 1000).toFixed(0);
    console.log(`  [${elapsed}s] CF=${s.cf}  "${s.title}"`);
    if (!s.cf) return true;
  }
  return false;
}

// ═══════════════════════════════════════════════════════════
// YouTube 字幕 → PDF
// ═══════════════════════════════════════════════════════════
async function processYoutube(url, title, num, outPath) {
  const tmpVtt = path.join(
    (process.env.TEMP || process.env.TMP || 'C:/Windows/Temp'),
    `yt_${num}.vtt`
  );
  const tmpHtml = tmpVtt.replace('.vtt', '.html');
  const stem = outPath.replace(/\.pdf$/, '');

  // 1. yt-dlp 下载字幕
  console.log('  yt-dlp 下载字幕...');
  const dlResult = spawnSync('yt-dlp', [
    '--write-auto-subs', '--write-subs',
    '--sub-langs', 'en',
    '--skip-download',
    '--output', stem,
    url,
  ], { encoding: 'utf-8', timeout: 60000 });

  // 找到下载的 .vtt 文件（yt-dlp 会加 .en.vtt 后缀）
  const vttPath = stem + '.en.vtt';
  if (!fs.existsSync(vttPath)) {
    console.error('  FAIL: 字幕文件未生成，yt-dlp stderr:', dlResult.stderr?.substring(0, 200));
    // fallback：打印 YouTube 页面
    console.log('  Fallback: 打印 YouTube 页面');
    const { targetId, sessionId } = await openTab(url);
    try {
      await navigateAndWait(sessionId, url);
      await sleep(3000);
      const size = await printToPdf(sessionId, outPath);
      console.log(`  OK (页面): ${(size/1024).toFixed(0)} KB\n`);
    } finally {
      await cdp('Target.closeTarget', { targetId }).catch(()=>{});
    }
    return;
  }

  // 2. 解析 VTT → 纯文本段落
  const vtt = fs.readFileSync(vttPath, 'utf-8').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const seen = new Set();
  const lines = [];
  for (const line of vtt.split('\n')) {
    if (!line || line.includes('-->') || line.startsWith('WEBVTT') ||
        line.startsWith('Kind:') || line.startsWith('Language:') || /^\d+$/.test(line)) continue;
    const clean = line.replace(/<[^>]+>/g, '').trim();
    if (clean && !seen.has(clean)) { seen.add(clean); lines.push(clean); }
  }

  // 每 60 行合并成一段
  const paragraphs = [];
  for (let i = 0; i < lines.length; i += 60) {
    paragraphs.push(lines.slice(i, i + 60).join(' '));
  }

  // 3. 生成 HTML
  function esc(s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>body{font-family:Georgia,serif;max-width:760px;margin:40px auto;font-size:13px;line-height:1.9;color:#222;}
h1{font-size:16px;border-bottom:1px solid #ccc;padding-bottom:8px;}
.meta{color:#666;font-size:11px;margin-bottom:24px;}p{margin:0 0 14px;text-align:justify;}</style>
</head><body>
<h1>${esc(title)}</h1>
<div class="meta">YouTube: ${esc(url)}<br>Auto-generated transcript (English) · ${paragraphs.length} segments</div>
${paragraphs.map(p => `<p>${esc(p)}</p>`).join('\n')}
</body></html>`;

  fs.writeFileSync(tmpHtml, html, 'utf-8');

  // 4. CDP 打印 PDF
  const { targetId, sessionId } = await openTab('about:blank');
  try {
    const fileUrl = 'file:///' + tmpHtml.replace(/\\/g, '/');
    const loadP = once('Page.loadEventFired', sessionId);
    await cdp('Page.navigate', { url: fileUrl }, sessionId);
    await Promise.race([loadP, sleep(5000)]);
    await sleep(300);
    const size = await printToPdf(sessionId, outPath);
    console.log(`  OK (字幕 PDF): ${(size/1024).toFixed(0)} KB, ${paragraphs.length} 段\n`);
  } finally {
    await cdp('Target.closeTarget', { targetId }).catch(()=>{});
    try { fs.unlinkSync(tmpHtml); } catch {}
    // 保留 .vtt 文件备用（可选删除）
  }
}

// ═══════════════════════════════════════════════════════════
// PubMed → PMC / 出版商 PDF 提取
// 策略：
//   1. 找 PMC 链接 → 进 PMC → 找 /pdf/ 链接 → in-page fetch 下载
//   2. 无 PMC → 找出版商全文链接 → 进出版商页 → 找 PDF 按钮 → in-page fetch 下载
//   3. 下载后验证魔数 %PDF，不合格（付费墙）→ printToPDF 当前 landing page
// ═══════════════════════════════════════════════════════════
async function processPubmed(url, title, num, outPath) {
  const { targetId, sessionId } = await openTab('about:blank');
  try {
    // 1. 加载 PubMed 摘要页
    await navigateAndWait(sessionId, url);
    await sleep(2000);

    // 2. 优先找 PMC 链接（免费全文）
    const pmcRes = await cdp('Runtime.evaluate', {
      expression: `(() => {
        const a = document.querySelector('a[href*="pmc.ncbi.nlm.nih.gov/articles/pmid/"]');
        return a ? a.href : '';
      })()`,
      returnByValue: true,
    }, sessionId);
    const pmcUrl = pmcRes?.result?.value || '';

    if (pmcUrl) {
      console.log(`  发现 PMC 链接，导航至 PMC...`);
      await navigateAndWait(sessionId, pmcUrl, 20000);
      await sleep(4000);
      const pdfRes = await cdp('Runtime.evaluate', {
        expression: `(() => {
          const a = Array.from(document.querySelectorAll('a[href]'))
            .find(el => el.href.includes('/pdf/') && !el.href.toLowerCase().includes('suppl'));
          return a ? a.href : '';
        })()`,
        returnByValue: true,
      }, sessionId);
      const pdfUrl = pdfRes?.result?.value || '';
      if (pdfUrl) {
        console.log(`  PMC PDF: ${pdfUrl.substring(0, 80)}`);
        const size = await downloadViaInPageFetch(sessionId, pdfUrl, outPath);
        const magic = Buffer.alloc(4);
        const fd = fs.openSync(outPath, 'r'); fs.readSync(fd, magic, 0, 4, 0); fs.closeSync(fd);
        if (magic[0] === 0x25 && magic[1] === 0x50 && magic[2] === 0x44 && magic[3] === 0x46) {
          console.log(`  OK (PMC PDF): ${(size / 1024).toFixed(0)} KB\n`);
          return;
        }
        fs.unlinkSync(outPath);
        console.log('  PMC PDF 验证失败，尝试出版商链接...');
      }
      // 导回 PubMed 页继续找出版商链接
      await navigateAndWait(sessionId, url);
      await sleep(1000);
    }

    // 3. 找出版商全文链接
    const pubRes = await cdp('Runtime.evaluate', {
      expression: `(() => {
        const links = Array.from(document.querySelectorAll('.full-text-links-list a, a.link-item'));
        const pub = links.find(a => {
          const h = a.href || '';
          return h.startsWith('http') &&
            !h.includes('pmc.ncbi') && !h.includes('pubmed.ncbi') &&
            !h.includes('#') && !h.includes('account.ncbi');
        });
        return pub ? pub.href : '';
      })()`,
      returnByValue: true,
    }, sessionId);
    const publisherUrl = pubRes?.result?.value || '';

    if (!publisherUrl) {
      console.log('  无出版商链接，打印 PubMed 摘要页...');
      const size = await printToPdf(sessionId, outPath);
      console.log(`  OK (PubMed 摘要): ${(size / 1024).toFixed(0)} KB\n`);
      return;
    }

    console.log(`  出版商链接: ${publisherUrl.substring(0, 80)}`);
    await navigateAndWait(sessionId, publisherUrl, 20000);
    await sleep(5000);

    // Cloudflare 检测
    const cfCheck = await cdp('Runtime.evaluate', {
      expression: `document.title === 'Just a moment...' || !!document.querySelector('#challenge-form')`,
      returnByValue: true,
    }, sessionId);
    if (cfCheck?.result?.value === true) {
      console.log('  检测到 Cloudflare，等待...');
      await waitForCloudflare(sessionId);
      await sleep(2000);
    }

    // 4. 在出版商页面找 PDF 下载按钮（排除补充材料）
    const btnRes = await cdp('Runtime.evaluate', {
      expression: `(() => {
        const all = Array.from(document.querySelectorAll('a[href]'));
        const candidate = all.find(a => {
          const href = a.href.toLowerCase();
          const text = (a.innerText || '').toLowerCase().trim();
          const looksLikePdf = href.includes('/pdf') || href.endsWith('.pdf');
          const isSuppl = href.includes('suppl') || href.includes('supplement');
          const hasPdfText = text.includes('pdf') || text === 'download';
          return looksLikePdf && !isSuppl && hasPdfText;
        });
        return candidate ? candidate.href : '';
      })()`,
      returnByValue: true,
    }, sessionId);
    const pdfUrl = btnRes?.result?.value || '';

    if (pdfUrl) {
      console.log(`  PDF 链接: ${pdfUrl.substring(0, 80)}`);
      try {
        const size = await downloadViaInPageFetch(sessionId, pdfUrl, outPath);
        const magic = Buffer.alloc(4);
        const fd = fs.openSync(outPath, 'r'); fs.readSync(fd, magic, 0, 4, 0); fs.closeSync(fd);
        if (magic[0] === 0x25 && magic[1] === 0x50 && magic[2] === 0x44 && magic[3] === 0x46) {
          console.log(`  OK (出版商 PDF): ${(size / 1024).toFixed(0)} KB\n`);
          return;
        }
        fs.unlinkSync(outPath);
        console.log('  PDF 非有效文件（付费墙），打印出版商 landing page...');
      } catch (e) {
        if (fs.existsSync(outPath)) fs.unlinkSync(outPath);
        console.log(`  PDF 获取失败: ${e.message}，打印出版商 landing page...`);
      }
    } else {
      console.log('  未找到 PDF 下载按钮，打印出版商 landing page...');
    }

    // 5. 回退：打印当前出版商页面（付费墙 landing page）
    const size = await printToPdf(sessionId, outPath);
    console.log(`  OK (landing page): ${(size / 1024).toFixed(0)} KB\n`);

  } catch (e) {
    console.error(`  FAIL: ${e.message}\n`);
  } finally {
    await cdp('Target.closeTarget', { targetId }).catch(() => {});
  }
}

// ═══════════════════════════════════════════════════════════
// 单条目处理
// ═══════════════════════════════════════════════════════════
async function processItem(raw, index, total) {
  const { num, title, url } = parseEntry(raw);
  const displayNum = num || (START_NUM + index);
  const type = detectType(url);
  const filename = sanitizeFilename(`${displayNum}_${title}`) + '.pdf';
  const outPath = path.join(OUT_DIR, filename);

  console.log(`[${displayNum}/${END_NUM}] ${title.substring(0, 70)}`);
  console.log(`  类型: ${type}  URL: ${url?.substring(0, 80)}`);

  if (!url) {
    console.log('  SKIP: 未找到 URL\n');
    return;
  }

  // ── YouTube → yt-dlp 字幕转 PDF ─────────────────────────
  if (type === 'youtube') {
    await processYoutube(url, title, displayNum, outPath);
    return;
  }

  // ── PubMed → PMC / 出版商 PDF ───────────────────────────
  if (type === 'pubmed') {
    await processPubmed(url, title, displayNum, outPath);
    return;
  }

  // ── 直接下载（PDF / Word 文档）──────────────────────────
  if (type === 'direct_pdf') {
    try {
      // 先下载到临时文件，检测实际类型
      const tmpDl = outPath + '.tmp_dl';
      const size = await downloadPdf(url, tmpDl);
      // 读取文件头魔数：OLE (.doc) = D0CF11E0，ZIP (.docx/xlsx) = 504B0304
      const buf4 = Buffer.alloc(4);
      const fd = fs.openSync(tmpDl, 'r');
      fs.readSync(fd, buf4, 0, 4, 0);
      fs.closeSync(fd);
      const magic = buf4.toString('hex').toUpperCase();
      const isDoc  = magic === 'D0CF11E0';  // Word 97-2003 .doc
      const isDocx = magic === '504B0304' && url.toLowerCase().includes('.docx'); // .docx (ZIP)
      if (isDoc || isDocx) {
        // 重命名为 .doc/.docx，调 Word 转 PDF
        const ext = isDocx ? '.docx' : '.doc';
        const docTmp = outPath.replace(/\.pdf$/, ext);
        fs.renameSync(tmpDl, docTmp);
        console.log(`  检测到 Word 文件 (${ext})，调用 Word COM 转换...`);
        convertWordToPdf(path.resolve(docTmp), path.resolve(outPath));
        fs.unlinkSync(docTmp);
        const pdfSize = fs.statSync(outPath).size;
        console.log(`  OK (Word→PDF): ${(pdfSize / 1024).toFixed(0)} KB\n`);
      } else {
        // 普通 PDF，直接重命名
        fs.renameSync(tmpDl, outPath);
        console.log(`  OK (直接下载): ${(size / 1024).toFixed(0)} KB\n`);
      }
    } catch (e) {
      console.error(`  FAIL: ${e.message}\n`);
    }
    return;
  }

  // ── CDP 处理（html / researchgate / pmc / frontiersin）──────────────────────
  const { targetId, sessionId } = await openTab('about:blank');
  try {
    if (type === 'pmc' || type === 'frontiersin' || type === 'researchgate') {
      // 通用 landing-page → PDF 提取逻辑
      // 1. 加载页面
      const loadP = once('Page.loadEventFired', sessionId);
      await cdp('Page.navigate', { url }, sessionId);
      await Promise.race([loadP, sleep(20000)]);

      // 2. 等待 Cloudflare（ResearchGate / Frontiersin 有可能触发）
      if (type === 'researchgate') {
        const passed = await waitForCloudflare(sessionId);
        if (!passed) console.log('  警告: Cloudflare 未自动通过');
      }
      await sleep(3000);

      // 3. 提取 PDF 下载链接
      const linkRes = await cdp('Runtime.evaluate', {
        expression: `(() => {
          // ResearchGate: fulltext PDF link
          if (location.hostname.includes('researchgate')) {
            const a = document.querySelector('a[href*="/fulltext/"][href$=".pdf"]');
            if (a) return a.href;
          }
          // Frontiersin: CDN PDF link (public-pages-files domain or /pdf path)
          if (location.hostname.includes('frontiersin')) {
            const a = document.querySelector('a[href*="frontiersin.org"][href$="pdf"], a[href*="public-pages"][href$="pdf"]');
            if (a) return a.href;
          }
          // PMC: /pdf/*.pdf link
          if (location.hostname.includes('ncbi')) {
            const a = document.querySelector('a[href*="/pdf/"][href$=".pdf"]');
            if (a) return a.href;
          }
          // Generic fallback: any link with .pdf that looks like a download
          const all = Array.from(document.querySelectorAll('a[href]'));
          const dl = all.find(el =>
            el.href.endsWith('.pdf') &&
            (el.innerText||'').toLowerCase().includes('pdf') &&
            !el.href.includes('citation') && !el.href.includes('reference')
          );
          return dl ? dl.href : '';
        })()`,
        returnByValue: true,
      }, sessionId);
      const pdfUrl = linkRes?.result?.value || '';

      if (!pdfUrl) {
        console.log('  未找到全文 PDF 链接，打印 landing page');
        const size = await printToPdf(sessionId, outPath);
        console.log(`  OK (landing page): ${(size / 1024).toFixed(0)} KB\n`);
      } else if (type === 'frontiersin') {
        // Frontiersin CDN 无需 cookies，直接下载
        console.log(`  Frontiersin CDN 直接下载...`);
        const size = await downloadPdf(pdfUrl, outPath);
        console.log(`  OK: ${(size / 1024).toFixed(0)} KB\n`);
      } else {
        // PMC / ResearchGate：需要浏览器 session，in-page fetch
        console.log(`  In-page fetch 下载 (${type})...`);
        const size = await downloadViaInPageFetch(sessionId, pdfUrl, outPath);
        console.log(`  OK: ${(size / 1024).toFixed(0)} KB\n`);
      }
      return;
    } else {
      await navigateAndWaitIdle(sessionId, url);
      await sleep(2000);
      // 检测 Cloudflare 拦截（对所有 html 类型页面）
      const cfCheck = await cdp('Runtime.evaluate', {
        expression: `document.title === 'Just a moment...' || !!document.querySelector('#challenge-form')`,
        returnByValue: true,
      }, sessionId);
      if (cfCheck?.result?.value === true) {
        console.log('  检测到 Cloudflare，等待自动通过...');
        const passed = await waitForCloudflare(sessionId);
        if (!passed) console.log('  警告: Cloudflare 未自动通过，尝试打印当前页面');
        await sleep(2000);
      }
      // ── 关闭 cookie / consent 弹窗 ──
      const cookieRes = await cdp('Runtime.evaluate', {
        expression: `(() => {
          const sels = [
            '.cb-enable',
            '#onetrust-accept-btn-handler',
            '.osano-cm-accept-all',
            '#ccc-recommended-settings',
            'button[class*="accept-cookies"]',
            '[data-testid="accept-button"]',
            'button.agree-button',
            '.fc-cta-consent',
          ];
          for (const s of sels) {
            const btn = document.querySelector(s);
            if (btn && btn.offsetHeight > 0) { btn.click(); return s; }
          }
          // 通用回退：在 cookie/consent 容器中找 Accept/OK/Agree 按钮
          const containers = document.querySelectorAll('[class*="cookie"], [class*="consent"], [id*="cookie"], [id*="consent"], [role="dialog"]');
          for (const c of containers) {
            if (c.offsetHeight === 0) continue;
            const btn = Array.from(c.querySelectorAll('button, a.btn, a[role="button"]'))
              .find(b => /accept|agree|ok|understand|got it|allow/i.test(b.innerText) && b.offsetHeight > 0);
            if (btn) { btn.click(); return 'generic:' + btn.innerText.trim().substring(0,30); }
          }
          return '';
        })()`,
        returnByValue: true,
      }, sessionId);
      const dismissed = cookieRes?.result?.value || '';
      if (dismissed) {
        console.log(`  关闭弹窗: ${dismissed}`);
        await sleep(1000);
      }
      // ── 多级 PDF URL 发现 ──
      const htmlBtnRes = await cdp('Runtime.evaluate', {
        expression: `(() => {
          // 1. meta 标签（学术出版商标准）
          const meta = document.querySelector('meta[name="citation_pdf_url"]');
          if (meta?.content) return meta.content;
          // 2. 可见 PDF 下载链接（含 /download/ 路径）
          const all = Array.from(document.querySelectorAll('a[href]'));
          const candidate = all.find(a => {
            const href = a.href.toLowerCase();
            const text = (a.innerText || '').toLowerCase().trim();
            const looksLikePdf = href.includes('/pdf') || href.endsWith('.pdf') || href.includes('/download/');
            const isSuppl = href.includes('suppl') || href.includes('supplement');
            const hasPdfText = text.includes('pdf') || text === 'download';
            return looksLikePdf && !isSuppl && hasPdfText;
          });
          if (candidate) return candidate.href;
          // 3. 从 DOI 构造 PDF URL（Atypon 等平台通用）
          const doi = document.querySelector('meta[name="citation_doi"]')?.content;
          if (doi) return location.origin + '/doi/pdf/' + doi;
          return '';
        })()`,
        returnByValue: true,
      }, sessionId);
      const htmlPdfUrl = htmlBtnRes?.result?.value || '';
      if (htmlPdfUrl) {
        console.log(`  PDF 链接: ${htmlPdfUrl.substring(0, 80)}`);
        try {
          const size = await downloadPdfViaNetworkResource(sessionId, htmlPdfUrl, outPath);
          const magic = Buffer.alloc(4);
          const fd = fs.openSync(outPath, 'r'); fs.readSync(fd, magic, 0, 4, 0); fs.closeSync(fd);
          if (magic[0] === 0x25 && magic[1] === 0x50 && magic[2] === 0x44 && magic[3] === 0x46) {
            console.log(`  OK (PDF): ${(size / 1024).toFixed(0)} KB\n`);
            return;
          }
          fs.unlinkSync(outPath);
          console.log('  PDF 验证失败（付费墙），打印页面...');
        } catch (e) {
          if (fs.existsSync(outPath)) fs.unlinkSync(outPath);
          console.log(`  PDF 获取失败: ${e.message}，打印页面...`);
        }
      }
      // 下载 PDF 前已导航走了的情况：回到原页面以便打印
      const curUrl = await cdp('Runtime.evaluate', { expression: 'location.href', returnByValue: true }, sessionId);
      if (curUrl?.result?.value !== url) {
        await navigateAndWaitIdle(sessionId, url);
        await sleep(2000);
      }
    }
    // 先尝试 Readability（文章页侧边栏/多栏布局时排版更佳）
    // 返回 false → 非文章页，回退到 printToPdf
    const rdSize = await printWithReadability(sessionId, outPath);
    if (rdSize !== false) {
      console.log(`  OK (Readability): ${(rdSize / 1024).toFixed(0)} KB\n`);
    } else {
      const size = await printToPdf(sessionId, outPath);
      console.log(`  OK: ${(size / 1024).toFixed(0)} KB\n`);
    }
  } catch (e) {
    console.error(`  FAIL: ${e.message}\n`);
  } finally {
    await cdp('Target.closeTarget', { targetId }).catch(() => {});
  }
}

// ═══════════════════════════════════════════════════════════
// 主流程
// ═══════════════════════════════════════════════════════════
async function main() {
  // 读取 Excel 条目
  const rows = readExcelRows(EXCEL_PATH, START_NUM, END_NUM);
  console.log(`读到 ${rows.length} 条条目\n`);

  if (rows.length === 0) {
    console.log('没有读到数据，请检查行范围和 Excel 格式。');
    process.exit(0);
  }

  // 判断是否需要 CDP（有 html 或 researchgate 类型）
  const needsCDP = rows.some(r => {
    const { url } = parseEntry(r);
    const t = detectType(url);
    return t !== 'direct_pdf';  // youtube/html/researchgate 都需要 CDP
  });

  if (needsCDP) {
    await connectCDP();
  }

  for (let i = 0; i < rows.length; i++) {
    await processItem(rows[i], i, rows.length);
  }

  console.log('═══════════════════════════════');
  console.log(`完成！文件保存在: ${OUT_DIR}`);
  console.log('═══════════════════════════════');

  if (needsCDP) {
    ws.close();
  }
  process.exit(0);
}

main().catch(e => {
  console.error('Fatal:', e);
  process.exit(1);
});
