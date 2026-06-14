// 设此 env 让 cdp-proxy.mjs 只导出 resolveConfig、不启动 server
process.env.CDP_PROXY_NO_MAIN = '1';
import assert from 'node:assert';
import os from 'node:os';
import path from 'node:path';
const { resolveConfig } = await import('./cdp-proxy.mjs');

// 1) 默认 env(无 override)→ 与现状一致
const d = resolveConfig({});
assert.strictEqual(d.proxyPort, 3456, 'default proxyPort');
assert.strictEqual(d.chromePort, 9222, 'default chromePort');
assert.strictEqual(d.tokenFile,
  path.join(os.homedir(), '.claude', 'cdp-proxy-token'), 'default tokenFile 路径不变');

// 2) 多实例 override
const m = resolveConfig({ CDP_PROXY_PORT: '3458', CDP_CHROME_PORT: '9224',
  CDP_USER_DATA_DIR: 'X:/inst-2', CDP_TOKEN_FILE: 'X:/tok-3458' });
assert.strictEqual(m.proxyPort, 3458);
assert.strictEqual(m.chromePort, 9224);
assert.strictEqual(m.userDataDir, 'X:/inst-2');
assert.strictEqual(m.tokenFile, 'X:/tok-3458');

// 3) 非默认端口但未显式给 token → 派生独立 token(防多实例互相覆盖)
const t = resolveConfig({ CDP_PROXY_PORT: '3459' });
assert.strictEqual(t.tokenFile,
  path.join(os.homedir(), '.claude', 'cdp-proxy-token-3459'), 'PORT≠3456 派生独立 token');

console.log('OK resolveConfig');
