# Pi for Office 中文技术文档

> [!NOTE]
> 本文档由 AI 辅助编写，用于中文用户/开发者深入了解 Pi for Office 的功能、用法与技术架构。如与英文文档不一致，请以英文版为准。

- 项目仓库：https://github.com/tmustier/pi4office
- 简要中文安装指南：[README.zh-CN.md](../README.zh-CN.md)
- 完整英文文档：[README.md](../README.md) 与 [docs/ 文档索引](./README.md)

---

## 目录

1. [项目概述](#1-项目概述)
2. [功能详解](#2-功能详解)
3. [技术架构](#3-技术架构)
4. [请求管线（Pipeline）](#4-请求管线pipeline)
5. [安装方式](#5-安装方式)
6. [使用指南](#6-使用指南)
7. [目录结构](#7-目录结构)
8. [开发与测试](#8-开发与测试)
9. [常见问题](#9-常见问题)

---

## 1. 项目概述

**Pi for Office** 是一款开源、多模型的 Microsoft Office AI 侧边栏加载项，支持 **Excel** 和 **Word**（另有 WPS 表格实验性支持）。它由 [Pi](https://pi.dev) 框架驱动——`pi-agent-core` 提供 agent 循环、`pi-ai` 提供多模型抽象、`pi-web-ui` 提供共享 UI 组件。

它本质上是一个**运行在 Office 内部的 AI 智能体**：能读取你的文档、直接修改内容、执行研究，模型由你选择。

**技术栈：**

| 层 | 技术 |
|---|---|
| 构建 | Vite 8（dev server + 生产打包） |
| UI | Lit web components |
| Agent 运行时 | `@earendil-works/pi-agent-core` |
| 模型抽象 | `@earendil-works/pi-ai` |
| 宿主 API | Office.js（Excel/Word）、WPS JSAPI |
| 认证 | OAuth PKCE + API Key + 自定义网关 |

**版本**：v0.10.0-pre（开发中）。

---

## 2. 功能详解

### 2.1 Excel 核心工具（17 个）

工具名与注册的**唯一事实来源**是 `src/tools/registry.ts`。所有 UI 渲染、输入人话化、系统提示词中的工具说明都从它派生。

| 工具 | 作用 | 类型 |
|---|---|---|
| `get_workbook_overview` | 结构蓝图——工作表、表头、命名区域、表格、图表、数据透视表 | 读 |
| `read_range` | 读单元格（紧凑 markdown / CSV / 含格式的详细模式） | 读 |
| `write_cells` | 写值/公式，带**覆盖保护**与自动校验 | 写 |
| `fill_formula` | 跨区域自动填充公式（相对引用自动调整） | 写 |
| `search_workbook` | 全簿搜索文本/值/公式引用 | 读 |
| `modify_structure` | 插入/删除行列、增删改名隐藏工作表 | 写（结构） |
| `format_cells` | 字体、颜色、数字格式、边框、命名样式 | 写 |
| `conditional_format` | 添加/清除条件格式 | 写 |
| `charts` | 创建/删除图表、获取图表图片 | 读写 |
| `trace_dependencies` | 公式溯源（上游引用 / 下游被引用） | 读 |
| `explain_formula` | 用通俗语言解释公式并引用单元格 | 读 |
| `view_settings` | 网格线、标题、冻结窗格、标签颜色、工作表可见性 | 读写 |
| `comments` | 批注：读/加/改/回复/解决/重开 | 读写 |
| `workbook_history` | 列出/恢复"保存间隙"自动备份 | 读写 |
| `instructions` | 持久化用户级与工作簿级规则 | 读写 |
| `conventions` | 格式化默认值（货币、负数、零、小数位） | 读写 |
| `skills` | 捆绑 Agent Skills 的读写与安装 | 读写 |

> 注：英文 README 表格写的是 16 个，实际注册表有 17 个（`charts` 未入表）。

### 2.2 Word 工具

| 工具 | 作用 |
|---|---|
| `get_document_outline` | 结构蓝图——节、标题、段落/字符数 |
| `read_document` | 读正文或选区，标题感知的 markdown 输出 |
| `insert_text` | 在光标/选区插入文本（替换/开头/结尾/前/后） |
| `search_document` | 全文搜索，带上下文与段落级结果 |
| `format_document` | 字体、字号、粗斜体、下划线、颜色、高亮、对齐 |
| `add_comment` | 为选中文本添加对话式批注 |

Word 与 Excel 共享：`instructions`、`conventions`、`skills`、`extensions_manager`、`execute_office_js`、Python 工具等。

### 2.3 通用 / 扩展工具

| 工具 | 作用 |
|---|---|
| `execute_office_js` / `execute_wps_js` | 结构化工具无法表达时直接调宿主 JSAPI（有审批门槛） |
| `python_run` | 在浏览器内跑 Python（Pyodide / WebAssembly），可处理 numpy/pandas/scipy |
| `python_transform_range` | 读范围进 Python 处理再写回（一次调用） |
| `libreoffice_convert` | 通过 LibreOffice 转换文档格式（PDF/DOCX/ODT 等） |
| `tmux` | 在用户机器上跑 shell 命令（需本地 tmux bridge） |
| `files` | 跨会话文件工作区（`notes/`、文档目录、`scratch/`、只读 `assistant-docs/`） |
| `extensions_manager` | 从聊天直接生成并安装侧边栏扩展 |
| `web_search` / `fetch_page` | 联网搜索（Jina/Serper/Tavily/Brave）+ 抓取页面 |
| `mcp` | 连接用户配置的 MCP 服务器，获得自定义工具 |

### 2.4 多模型支持

- **Anthropic**（Claude）—— API Key 或 OAuth
- **OpenAI** / **OpenAI Codex** —— API Key
- **Google Gemini** / Code Assist / Antigravity —— API Key 或 OAuth
- **GitHub Copilot** —— OAuth
- **自定义 OpenAI 兼容网关** —— `/settings` 配置端点 + 模型 + Key（如 DeepSeek、智谱 GLM、Ollama 本地模型）

对话中可随时 `/model` 切换模型；模型排序与版本评分逻辑在 `src/models/`。

### 2.5 会话管理

- 每个工作簿/文档可开**多个会话标签页**
- 自动保存/恢复、会话历史、`/resume` 续接
- 会话标题自动生成，可显式命名

### 2.6 自动上下文注入

每轮对话前，插件自动向模型注入：
- 工作簿**结构蓝图**（仅首次/结构变化时刷新）
- 工作区文件摘要
- 当前**选区**内容
- 自上次消息以来的**单元格改动**摘要

即你不需要描述"我在看什么"。

### 2.7 工作簿恢复与执行模式

- **恢复检查点**：每次写入前快照受影响单元格，写入恢复日志，出错可一键回滚；另有"保存间隙"自动备份（`workbook_history`）
- **执行模式**：
  - **Auto**（yolo）：变更立即执行，不逐项确认
  - **Confirm**（safe）：每次工作簿变更前询问用户

### 2.8 斜杠命令

`/model` `/login` `/settings` `/rules` `/extensions` `/tools` `/export` `/compact` `/new` `/resume` `/history` `/shortcuts` `/yolo` `/experimental` 等，注册表在 `src/commands/`。

### 2.9 扩展系统

- 可从聊天让 AI 生成并安装侧边栏扩展（`extensions_manager`）
- 扩展默认运行在 **iframe 沙箱**；内置/本地模块运行在宿主
- 有权限模型、远程 URL 白名单、Widget API v2

### 2.10 集成

- **Web Search**：Jina（默认）/ Serper / Tavily / Brave + `fetch_page`
- **MCP Gateway**：连接用户配置的 MCP 服务器
- **本地 bridge**：Python/LibreOffice bridge（默认 Pyodide，可切 native）、tmux bridge
- **CORS 代理**：OAuth 在 Office 内嵌浏览器被 CORS 拦截时使用

### 2.11 规则 / 约定 / 记忆

- **`instructions`**：持久化"所有文件"（本机私有）与"此文件"两级规则
- **`conventions`**：定义一次格式风格（货币符号、负数样式、小数位），AI 自动遵循
- **工作区记忆**：`files` 下的 `notes/` 与文档目录作为跨会话持久记忆

---

## 3. 技术架构

```
┌─────────────────────────────────────────────────────────┐
│  Office 宿主层 (host/)                                   │
│  OfficeHost / WpsHost / BrowserHost —— 检测宿主、onReady │
├─────────────────────────────────────────────────────────┤
│  Taskpane 应用层 (taskpane/)                             │
│  init/bootstrap · 会话管理 · tab 布局 · 上下文注入 · 状态栏│
├─────────────────────────────────────────────────────────┤
│  Agent 运行时 (pi-agent-core)                            │
│  工具循环 · 流式输出 · 状态管理                          │
├─────────────────────────────────────────────────────────┤
│  工具层 (tools/) — 17 个 Excel + Word + 通用工具         │
│  registry.ts 单一事实来源 · execution-policy · mutation  │
├─────────────────────────────────────────────────────────┤
│  上下文层 (context/) · 提示词 (prompt/) · 模型 (models/) │
│  blueprint 缓存 · change-tracker · system-prompt 构建    │
├─────────────────────────────────────────────────────────┤
│  认证层 (auth/) — OAuth · API Key · 代理校验 · 网关      │
├─────────────────────────────────────────────────────────┤
│  存储层 (storage/) — IndexedDB · provider 凭据           │
└─────────────────────────────────────────────────────────┘
```

### 3.1 关键设计模式

- **工具注册表作为单一事实来源**：`src/tools/registry.ts` 定义所有核心工具名与构造；UI 渲染器、人话化参数、提示词文档均从中派生。
- **Workbook 协调器**：核心不变量是**并行推理、串行化写入**——同一工作簿的写操作排队，防止多会话标签页并发写冲突（`src/workbook/coordinator.ts`）。
- **执行策略**：每个工具分类为 `read` 或 `mutate`，写入再分为 `content` / `structure`，决定加锁、检查点与蓝图失效行为（`src/tools/execution-policy.ts`）。
- **恢复检查点**：写入前快照受影响单元格，支持一键回滚（`src/workbook/recovery/`）。
- **扩展沙箱**：不可信扩展（内联代码、远程 URL）默认跑在 iframe 沙箱；内置/本地模块跑在宿主。
- **提示缓存友好设计**：系统提示词前缀保持静态稳定，易变状态一律走消息尾部注入，避免缓存键漂移。

### 3.2 工具包装管线

每个工具在注册进 Agent 前经过三层包装（`src/taskpane/init.ts`）：

```
裸工具
  → withWorkbookCoordinator   （串行写入 + Confirm 审批 + mutation 观察者→蓝图失效）
  → withConnectionPreflight   （web_search/mcp 等连接前置检查，未配置时引导设置）
  → applyToolOutputTruncation （按模型上下文窗口截断工具输出）
```

---

## 4. 请求管线（Pipeline）

### 4.1 启动流程

```
taskpane.ts → bootstrap → 检测宿主 (Office/WPS/Browser) → Office.onReady
  → 初始化 IndexedDB 存储 → 恢复会话标签页
  → new Agent({ systemPrompt, model, tools, transformContext, streamFn })
```

### 4.2 每轮用户消息的完整链路

1. **上下文注入**（`transformContext` = `createContextInjector`）：把 `[Auto-context]` 段插到最后一条用户消息之前——工作簿蓝图（按 revision 缓存）、工作区文件摘要、选区、改动摘要。
2. **System prompt 构建**：按 Excel/Word 分身份，含规则、执行模式、活跃集成/连接状态、本地服务探测、可用 skills、工具说明、工作流指引、格式约定。
3. **模型推理**：`pi-ai` 多 provider 流式输出。
4. **工具执行**（关键不变量）：
   - 协调器串行化写入；Confirm 模式先弹审批
   - 写入前快照 → 恢复日志
   - 写完后仅 `structure` 级改动失效蓝图（触发下轮重新注入）
5. **结果渲染**：工具结果以 `content`（人类可读）+ `details`（稳定机器元数据）双通道返回。

---

## 5. 安装方式

Pi for Office 提供四种安装路径，按网络环境与工具偏好选择：

### 5.1 方式一：Windows 一键安装器（推荐）

1. 从 [GitHub Releases](https://github.com/tmustier/pi4office/releases/latest) 下载 `pi4office-setup.exe`
2. 运行安装器——自动生成证书、注册加载项、创建开始菜单快捷方式
3. 从开始菜单启动 **Pi for Office Server**
4. Excel → **开始(Home) → 加载项(Add-ins) → Pi for Office**

安装器内置本地 HTTPS 服务与 CORS 代理，无需 Node.js 或额外配置。

### 5.2 方式二：手动侧载（托管版，零本地进程）

1. 下载 [manifest.prod.xml](https://pi4office.vercel.app/manifest.prod.xml)
2. 按 [docs/install.md](./install.md) 加入 Excel/Word
3. 功能区点 **Open Pi**

> 插件本体托管在 Vercel，无本地服务。**注意**：在某些网络（如中国大陆）下 `*.vercel.app` 可能被 DNS 污染而无法访问；跑的是已部署版本而非本地代码。

### 5.3 方式三：传统共享文件夹目录（Windows）

1. 右键文件夹 → **属性 → 共享 → 共享**，记下网络路径（如 `\\YourPC\Addins`）
2. Excel：**文件 → 选项 → 信任中心 → 信任中心设置 → 受信任的加载项目录**
3. 把网络路径加为 **目录 URL**，勾选**在菜单中显示**，重启 Excel
4. 把 `manifest.prod.xml`（本地模式则为 `manifest.xml`）放进共享文件夹
5. **开始 → 加载项 → 高级 → 共享文件夹 → Pi for Office**

### 5.4 方式四：本地生产构建服务（无 Vercel、无 npm run dev）

适用于托管 URL 被屏蔽的网络（如中国大陆），把生产构建跑在本机：

```bash
git clone https://github.com/tmustier/pi4office.git
cd pi4office
npm install
mkcert -install          # 生成 HTTPS 证书（一次性）
mkcert localhost
mv localhost-key.pem key.pem
mv localhost.pem cert.pem

npm run build            # 生产构建 → dist/
npm run serve:dist       # HTTPS 服务 https://localhost:3141（保持运行）
```

然后侧载仓库根目录的 **`manifest.xml`**（指向 localhost:3141），而非 `manifest.prod.xml`。注意：无热更新，改代码需重新 `npm run build`。

### 5.5 方式五：开发者模式（本地 Vite）

```bash
npm install
mkcert -install && mkcert localhost && mv localhost-key.pem key.pem && mv localhost.pem cert.pem
npm run use        # 起 HTTPS dev server（端口 3141）并侧载到 Excel/Word
```

---

## 6. 使用指南

### 6.1 连接模型

**API Key（推荐，无需代理）**：`/login` → 展开服务商 → 粘贴 Key → Save。

**自定义 OpenAI 兼容网关**：`/settings` → Custom OpenAI-compatible gateways → 填 Endpoint / Model / Key（本地服务 Key 可留空）→ `/model` 选择。

| 服务商 | Endpoint | 模型 ID 示例 |
|---|---|---|
| DeepSeek | `https://api.deepseek.com` | `deepseek-chat`、`deepseek-reasoner` |
| 智谱 GLM | `https://open.bigmodel.cn/api/paas/v4` | `glm-4.6` 等 |
| Ollama（本地） | `http://localhost:11434/v1` | 本地已下载的模型 |

**OAuth 登录**：`/login` → **Login with …** → 浏览器完成登录 → 返回 Excel。浏览器最后跳到"无法访问此网站"是正常现象，复制地址栏 URL 粘贴回提示框即可。

**OAuth 报 CORS 错误**：本机跑 `npx pi4office-proxy`（默认 `https://localhost:3003`），然后在 `/settings → Proxy` 启用并填入地址，重试。

### 6.2 对话使用示例

- 问：`有哪些工作表？` / `Outline this document`
- 让它改：`把 A 列求和放在 E15`、`给表头加粗上色`
- 分析：`总结我选中的区域`

插件会自动注入结构蓝图、当前选区与最近改动，无需手动描述。

### 6.3 常用命令速查

| 命令 | 作用 |
|---|---|
| `/model` | 切换模型 |
| `/new` `/resume` `/history` | 新建/恢复/查看会话 |
| `/rules` | 查看/编辑持久规则 |
| `/tools` | 管理 Web Search + MCP 连接 |
| `/extensions` | 安装/管理扩展 |
| `/export` `/compact` | 导出对话 / 压缩超长会话 |
| `/yolo` | 切 Auto 模式（免确认） |
| `/experimental` | 管理实验性功能与本地 bridge |

### 6.4 安全使用要点

- 写操作有**覆盖保护**：目标已有数据会拦截并询问，需显式允许 `allow_overwrite`
- Confirm 模式（`/settings`）每次变更前询问；Auto 模式（`/yolo`）直接执行
- 误删可经恢复检查点一键回滚
- 扩展默认在 iframe 沙箱运行；远程 URL 扩展需显式开启

---

## 7. 目录结构

顶层与 `src/` 结构见 [架构说明](#3-技术架构) 与英文 README 的 [Source layout](../README.md#source-layout)。核心模块速查：

| 目录 | 职责 |
|---|---|
| `src/taskpane/` | 应用初始化、会话管理、上下文注入、tab 布局 |
| `src/tools/` | 17 个 Excel + Word + 通用工具、执行策略、检查点 |
| `src/commands/` | 斜杠命令注册 + builtins |
| `src/context/` | 工作簿蓝图缓存、变更追踪、选区 |
| `src/prompt/` | 系统提示词构建 |
| `src/auth/` | OAuth 提供方、代理校验、凭据恢复 |
| `src/models/` | 模型排序 + 版本评分 |
| `src/ui/` | 侧边栏组件、工具渲染、主题 |
| `src/extensions/` | 扩展商店、沙箱运行时、权限 |
| `src/workbook/` | 工作簿身份、协调器、恢复日志 |
| `src/host/` | Office/WPS/Browser 宿主抽象 |
| `scripts/` | dev 脚本、bridge server、check 脚本 |
| `pkg/` | 可发布 CLI 包（proxy / python-bridge / tmux-bridge / installer） |
| `tests/` | 单元 + 安全测试（~50 文件） |

---

## 8. 开发与测试

### 8.1 常用命令

| 命令 | 作用 |
|---|---|
| `npm run dev` | Vite dev server（HTTPS，端口 3141） |
| `npm run use` | 起 dev server 并侧载 |
| `npm run build` | 生产构建 → `dist/` |
| `npm run check` | lint + typecheck + CSS 主题检查 |
| `npm run test` | 全部测试（models + context + security + manifest） |
| `npm run test:context` | 工具/上下文/会话/扩展测试 |
| `npm run test:security` | 代理/CORS/沙箱/OAuth 安全测试 |
| `npm run proxy:https` | CORS 代理（OAuth 用） |

### 8.2 测试套件

| 套件 | 覆盖 |
|---|---|
| `test:models` | 模型排序、运行时模型对账 |
| `test:context` | 工具、上下文注入、会话、扩展、集成 |
| `test:security` | 代理目标策略、CORS、沙箱、OAuth、CSP |
| `test:manifest` | 清单生成 |

### 8.3 代码规范要点

- 代码注释、标识符仅使用英文
- 工具名唯一事实来源在 `registry.ts`，新增工具需同步更新 UI 渲染、人话化参数、上下文披露、提示词
- 用户可见 UI 字符串走 `t()`（i18n），agent 面向字符串不进 i18n
- 不做 `// @ts-ignore`、不引入不必要的 Node-only 依赖（保持 WebView 兼容）

---

## 9. 常见问题

- **加载项列表里看不到 Pi** —— 重启 Excel；确认上传的是 `manifest.prod.xml`（非 localhost 开发版清单）
- **侧边栏打开但空白** —— 网络可能屏蔽托管 URL，见[方式四：本地生产构建](#54-方式四本地生产构建服务无-vercel无-npm-run-dev)
- **OAuth 登录失败** —— 确认代理运行且在 `/settings` 填了正确的 HTTPS 地址；可改用 API Key
- **如何更新** —— 大多数更新自动生效（关闭重开侧边栏）；清单变更需重新上传 `manifest.prod.xml`
- **WPS 支持** —— 实验性，见 [docs/wps-support.md](./wps-support.md)

---

## 许可证

[MIT](../LICENSE) © Thomas Mustier
