#!/usr/bin/env node
// 环境检查 + 确保 CDP Proxy 就绪（跨平台，替代 check-deps.sh）
//
// Chrome 实例由 cdp-proxy.mjs 自己负责生命周期：
//   - 9222 已有 Chrome → 复用
//   - 9222 没有 → 启动独立 "Debug Data" 实例（固定 user-data-dir，保留登录态）
// 本脚本不再尝试发现用户日常 Chrome——历史上读 DevToolsActivePort 会找到用户主浏览器，
// 但 proxy 只看 9222，导致 "chrome: ok (port 57289)" 后 proxy 实际连不上。

import { spawn, execSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const PROXY_SCRIPT = path.join(ROOT, 'scripts', 'cdp-proxy.mjs');
const PROXY_PORT = Number(process.env.CDP_PROXY_PORT || 3456);
const PROXY_LOG = path.join(os.tmpdir(), 'cdp-proxy.log');

// --- Node.js 版本检查 ---

function checkNode() {
  const major = Number(process.versions.node.split('.')[0]);
  const version = `v${process.versions.node}`;
  if (major >= 22) {
    console.log(`node: ok (${version})`);
  } else {
    console.log(`node: warn (${version}, 建议升级到 22+)`);
  }
}

// --- HTTP / 进程工具 ---

function httpGetJson(url, timeoutMs = 3000) {
  return fetch(url, { signal: AbortSignal.timeout(timeoutMs) })
    .then(async (res) => {
      try { return JSON.parse(await res.text()); } catch { return null; }
    })
    .catch(() => null);
}

function findPidOnPort(port) {
  try {
    if (os.platform() === 'win32') {
      const out = execSync('netstat -ano', { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] });
      for (const line of out.split(/\r?\n/)) {
        if (!line.includes('LISTENING')) continue;
        const parts = line.trim().split(/\s+/);
        const local = parts[1];
        if (!local || !local.includes(':')) continue;
        // 精确匹配端口尾段，避免 13456 误中 3456
        if (local.split(':').pop() === String(port)) {
          const pid = parseInt(parts[parts.length - 1], 10);
          if (pid > 0) return pid;
        }
      }
    } else {
      try {
        const out = execSync(`lsof -ti:${port} -sTCP:LISTEN`, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] });
        const pid = parseInt(out.trim().split('\n')[0], 10);
        if (pid > 0) return pid;
      } catch { /* lsof 不存在或无匹配 */ }
    }
  } catch { /* netstat 偶发失败 */ }
  return 0;
}

function killPid(pid) {
  if (!pid || pid <= 0) return;
  try {
    if (os.platform() === 'win32') {
      execSync(`taskkill /F /PID ${pid}`, { stdio: 'ignore' });
    } else {
      process.kill(pid, 'SIGKILL');
    }
  } catch { /* 已死或无权限 */ }
}

function killStaleProxy() {
  const pid = findPidOnPort(PROXY_PORT);
  if (pid > 0) {
    killPid(pid);
    return pid;
  }
  return 0;
}

// --- CDP Proxy 启动与等待 ---

function startProxyDetached() {
  const logFd = fs.openSync(PROXY_LOG, 'a');
  const child = spawn(process.execPath, [PROXY_SCRIPT], {
    detached: true,
    stdio: ['ignore', logFd, logFd],
    ...(os.platform() === 'win32' ? { windowsHide: true } : {}),
  });
  child.unref();
  fs.closeSync(logFd);
}

async function waitForReady(maxSeconds = 25) {
  const healthUrl = `http://127.0.0.1:${PROXY_PORT}/health`;
  for (let i = 1; i <= maxSeconds; i++) {
    const h = await httpGetJson(healthUrl, 5000);
    if (h?.connected === true) return h;
    if (i === 3) {
      console.log('  (Chrome 首次启动较慢；若弹出调试授权对话框请点「允许」)');
    }
    await new Promise((r) => setTimeout(r, 1000));
  }
  return null;
}

// 处理三种状态：
//   1) proxy 健康（connected=true）          → 直接返回
//   2) proxy 在跑但未连接 Chrome（残留状态） → 杀掉重启（自愈）
//   3) proxy 未运行                          → 启动
async function ensureProxy() {
  const healthUrl = `http://127.0.0.1:${PROXY_PORT}/health`;
  const initial = await httpGetJson(healthUrl);

  if (initial?.connected === true) {
    console.log(`proxy: ready (chromePort=${initial.chromePort}, sessions=${initial.sessions})`);
    return true;
  }

  if (initial && initial.status === 'ok') {
    // proxy 在跑但 connected != true（如 chromePort=null）→ 残留实例无法自愈，必须杀
    const killedPid = killStaleProxy();
    if (killedPid) {
      console.log(`proxy: stale (running but disconnected, killed pid=${killedPid}), restarting...`);
      await new Promise((r) => setTimeout(r, 1000));
    } else {
      console.log('proxy: stale but pid not found, attempting restart anyway...');
    }
  } else {
    console.log('proxy: not running, starting...');
  }

  startProxyDetached();
  // 给进程时间起来
  await new Promise((r) => setTimeout(r, 1500));

  const ready = await waitForReady(25);
  if (ready) {
    console.log(`proxy: ready (chromePort=${ready.chromePort}, sessions=${ready.sessions})`);
    return true;
  }

  console.log('proxy: failed (timeout waiting for Chrome connection)');
  console.log(`  日志：${PROXY_LOG}`);
  return false;
}

// --- main ---

async function main() {
  checkNode();
  console.log('chrome: managed by proxy (auto-launches Debug Data instance on demand)');

  const ok = await ensureProxy();
  if (!ok) process.exit(1);

  // 列出已有站点经验
  const patternsDir = path.join(ROOT, 'references', 'site-patterns');
  try {
    const sites = fs.readdirSync(patternsDir)
      .filter(f => f.endsWith('.md'))
      .map(f => f.replace(/\.md$/, ''));
    if (sites.length) {
      console.log(`\nsite-patterns: ${sites.join(', ')}`);
    }
  } catch { /* 目录可能不存在 */ }
}

await main();
