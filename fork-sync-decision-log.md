# Fork Sync 决策日志（长期保留）

> 本文件记录本 fork（`yekeyekegh/web-access-fork`）每次与上游（`eze-is/web-access`）同步时的
> **同步前状态 / 决策与缘由 / 采取的动作 / 最终结果 / 涉及的 commit**。
>
> **约定（每次 sync 必读 + 必写）：**
> 1. 任何一次 sync 之前，**先读这份文件**，了解历史上「哪些上游提交被有意跳过、为什么」，避免重复推演或误合已被否决的改动。
> 2. 任何一次 sync 之后，**在「同步记录」区顶部新增一条**（按时间倒序，最新在上）。
> 3. 重点维护下方「**有意未合并的上游改动**」清单——这是本 fork 与上游的核心分歧，决定了未来合并的难度。
>
> 远端：
> - `origin` = https://github.com/yekeyekegh/web-access-fork.git （本 fork，PR/合并目标）
> - `upstream` = https://github.com/eze-is/web-access.git （上游）

---

## 有意未合并的上游改动（重要 · 长期维护）

| 上游提交 | 版本 | 内容 | 跳过原因 | 决定日期 |
|---|---|---|---|---|
| `918d933` | 2.5.2 | Edge/Chromium 支持 + 浏览器偏好持久化，抽出 `browser-discovery.mjs` 统一「发现」逻辑 | 与本 fork 核心适配「自动拉起独立 Debug Data Chrome 实例（零手动设置）」哲学冲突：上游只**发现**已手动开 remote debugging 的浏览器、自己不拉起。合并会丢失本 fork 的零设置体验。**代价已知**：未来若想要 Edge 支持，再合 core 会更费劲（两套发现逻辑需缝合）。 | 2026-06-14 |

> 同时：本 fork 保留上游已删除的 `scripts/fetch_refs.mjs`（独立 CLI，从 Excel 批量下载参考文献 PDF，上游 scope 外）——sync 时勿跟随上游删除。

---

## 本 fork 的核心适配（与上游的差异点，sync 时须保护）

- **自动拉起 Debug Data Chrome**：`cdp-proxy.mjs` 连不上 9222 时自动启动独立 Chrome 实例（固定 user-data-dir 保留登录态），无需手动开 remote debugging。上游无此功能。
- **token 鉴权**：除 `/health` 外所有 proxy 端点需 `?token=$TOKEN`（query）。上游无鉴权——故上游来源的文档示例都缺 token，移植时须补充说明。
- **`isAllowedUrl` 白名单**、`workWindowId`/newWindow 逻辑、截图路径限制等安全加固（commit `f981930`）。
- **`fetch_refs.mjs`**（见上）。

---

## 同步记录（按时间倒序，最新在上）

### 2026-06-14 — 同步至上游 v2.5.3（选择性移植，非完整合并）

**同步前状态**
- 本地版本：`plugin.json` = 2.4.3，`SKILL.md` frontmatter = 2.5.0（fork 内部版本号本就不一致）。
- 与上游的共同祖先（merge-base）：`3caf8f3`（2026-04-23 "Update README.md"）。
- 上游领先本 fork 3 个提交：
  - `8757c68`（2.5.1, 2026-05-15）plugin.json `skills` 字段 `"./"` → `["./"]`，修复 Claude Code 插件加载失败（Path escapes plugin directory，Closes #113/#108/#62/#59）。
  - `918d933`（2.5.2, 2026-05-15）Edge/Chromium + browser-discovery 重构。
  - `7af34af`（2.5.3, 2026-05-16）`/new`、`/navigate` 从 GET `?url=` 改为 POST body 传 URL，修复含 query 的目标 URL 被 `&` 切断的 bug（小红书 `xsec_token` 等，#111）。

**决策与缘由**
- 采用「**选择性 cherry-pick（方案 C）**」，不做完整 merge。
- 取 `8757c68`(2.5.1)：插件加载修复，刚需。
- 取 `7af34af`(2.5.3)：POST body 修复，正中本 fork 高频场景（小红书含 query URL），刚需。
- **跳过** `918d933`(2.5.2)：见上方「有意未合并」表——与自动拉起哲学冲突，舍弃会丢失零设置体验。
- 完整 merge 实测会在 `cdp-proxy.mjs`(6 块)、`check-deps.mjs`(8 块) 等产生大量 core 冲突，且会误删 `fetch_refs.mjs`，故放弃。

**采取的动作**
1. 在分支 `merge-upstream-2.5.3` 上操作（未直接动 main）。
2. 把 `/new`、`/navigate` 的 POST body 逻辑**移植**进本 fork 的 `cdp-proxy.mjs`（保留 `isAllowedUrl`、`workWindowId`、token 鉴权；旧 GET 写法返回 400 + 迁移提示）。
3. `plugin.json`：`skills` → `["./"]`，版本 2.4.3 → 2.5.3。`SKILL.md`：版本 2.5.0 → 2.5.3，API 示例改 POST。
4. `references/cdp-api.md` 改 POST；从上游拉入 `references/migration-2.5.3.md` 并补 fork 专属 token 说明。
5. 开 PR #1 → 经多 agent code review（code-reviewer / silent-failure-hunter / comment-analyzer）→ 走 receiving-code-review triage。
6. review 修复：`/navigate` 空 POST body 显式返回 400（此前 `isAllowedUrl("")` 返回 true 会静默导航空 URL）；README.md `/new` 旧 GET 改 POST；各文档补 token 说明。
7. 全程实跑 proxy + Chrome 验证：带 query 的 `/new`/`/navigate` URL 完整保留、空 body→400、自动拉起 Chrome 正常。
8. PR #1 以 merge commit 合并入 `main`，删除远端分支，本地/远端 main 同步一致。

**最终结果**
- 本地 = 远端 main = `6f0d5fc`。版本统一为 **2.5.3**。
- 2.5.1、2.5.3 已并入；2.5.2 有意保留 fork 自有实现。

**涉及的 commit**
- `efd3cf9`（2026-06-14）`feat: 移植 upstream v2.5.3 POST body 修复 + 插件加载修复`
  - cdp-proxy.mjs `/new`、`/navigate` 改 POST body（保留 isAllowedUrl/workWindowId/token）；plugin.json `skills`→`["./"]` + 版本 2.5.3；SKILL.md/cdp-api.md 示例改 POST；拉入 migration-2.5.3.md 并补 token 说明。未合 browser-discovery/Edge 重构与 check-deps 改写；保留 fetch_refs.mjs。
- `8b3a0b0`（2026-06-14）`fix(review): /navigate 空 body 拒绝 + 文档 token 说明补全`
  - cdp-proxy.mjs `/navigate` 空 body→400；README.md `/new` 旧 GET 改 POST + token 说明；cdp-api.md 基础信息补鉴权说明；migration-2.5.3.md 转换表加 fork 脚注 + FAQ 补「无 token 先返回 403」。
- `6f0d5fc`（2026-06-14）`Merge pull request #1 from yekeyekegh/merge-upstream-2.5.3`（合并入 main）。

**对应上游提交**：`8757c68`(2.5.1)、`7af34af`(2.5.3) 已并入；`918d933`(2.5.2) 跳过。
