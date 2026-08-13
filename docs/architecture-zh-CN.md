# Pi for Office 整体架构深入文档（中文）

> 本文是 Pi for Office 的**整体架构深入文档**。它面向希望理解项目内部结构的开发者，按「分层架构 → 关键机制 → 安全与发布」的次序，逐层说明加载项的模块划分、核心不变量与关键技术决策。

| 项目 | 值 |
|---|---|
| 版本基线 | v0.10.0-pre（依赖 pi-agent-core / pi-ai 0.80.8） |
| 项目仓库 | https://github.com/tmustier/pi4office |
| 阅读对象 | 希望深入理解本项目内部结构的开发者 |
| 配套文档 | 功能/用法概览见 [guide-zh-CN.md](./guide-zh-CN.md)；安装见 [install.md](./install.md) |

---

## 目录

1. [文档说明](#1-文档说明)
2. [项目概述与定位](#2-项目概述与定位)
3. [总体架构](#3-总体架构)
4. [宿主抽象层（Office / WPS / Browser）](#4-宿主抽象层office--wps--browser)
5. [启动与引导流程](#5-启动与引导流程)
6. [工具系统](#6-工具系统)
7. [并发与一致性：工作簿协调器](#7-并发与一致性工作簿协调器)
8. [上下文与提示词管理](#8-上下文与提示词管理)
9. [模型抽象与认证](#9-模型抽象与认证)
10. [存储层](#10-存储层)
11. [会话管理](#11-会话管理)
12. [斜杠命令系统](#12-斜杠命令系统)
13. [扩展系统](#13-扩展系统)
14. [集成与本地桥接](#14-集成与本地桥接)
15. [安全模型](#15-安全模型)
16. [构建与发布](#16-构建与发布)
17. [WPS 支持（实验性）](#17-wps-支持实验性)
18. [目录结构与模块职责](#18-目录结构与模块职责)
19. [开发与测试](#19-开发与测试)
20. [核心不变量速查与参考](#20-核心不变量速查与参考)

---

## 1. 文档说明

本文是 Pi for Office 的**整体架构深入文档（中文）**。它面向希望理解项目内部结构的开发者，按「分层架构 → 关键机制 → 安全与发布」的次序，逐层说明加载项的模块划分、核心不变量与关键技术决策。

### 1.1 定位与关系

- 本项目英文 README（`README.md`）与 docs/ 索引负责「功能与使用」视角。
- [docs/guide-zh-CN.md](./guide-zh-CN.md) 提供「功能 / 用法 / 架构概览」，本文在其架构章节上做完整展开，不再重复安装与日常使用步骤。
- 文中引用的文件路径、符号名均对应 v0.10.0-pre 基线（依赖 pi-agent-core / pi-ai 0.80.8）。

### 1.2 阅读约定

- 「**核心不变量**」小节列出必须保持的设计约束，改动时应重点回归。
- 代码块为架构示意图或关键配置片段，非完整源码。
- 「单一事实来源」「并行推理、串行写入」等短语为项目内部术语，见对应章节解释。

---

## 2. 项目概述与定位

Pi for Office 是开源、多模型的 Microsoft Office AI 侧边栏加载项，支持 **Excel** 与 **Word**，另有 **WPS 表格**的实验性支持。它本质上是一个**运行在 Office 内部的 AI 智能体**：读取文档、直接修改内容、执行研究，模型由用户选择（Anthropic / OpenAI / Gemini / GitHub Copilot / 任意 OpenAI 兼容网关）。

### 2.1 技术栈

| 层 | 技术 | 职责 |
|---|---|---|
| 构建 | Vite（dev server + 生产打包） | HTTPS 本地服务、反向代理、浏览器适配 |
| UI | Lit web components + pi-web-ui | 侧边栏组件、消息/工具渲染、主题 |
| Agent 运行时 | `@earendil-works/pi-agent-core` | 工具循环、消息状态、流式输出、上下文变换 |
| 模型抽象 | `@earendil-works/pi-ai` | 多 provider 统一、模型发现、流式推理 |
| 宿主 API | Office.js（Excel/Word）、WPS JSAPI | 文档读写、选区、结构 |
| 认证 | OAuth PKCE + API Key + 自定义网关 | 多服务商登录与凭据管理 |
| 存储 | IndexedDB（数据库名 `pi4office`，版本 2） | 设置、凭据、会话、自定义网关、模型目录 |
| 本地进程 | Python/LibreOffice 桥、tmux 桥、CORS 代理 | Pyodide 之外的扩展能力与 OAuth 代理 |

### 2.2 设计目标

- 在 Office 宿主受限的 WebView 中运行完整 agent：无 Node 运行时，纯浏览器内推理。
- 多会话并行推理，但对同一工作簿的写入严格串行，保证数据一致性。
- 通过存储卫生与沙箱，把不可信代码（扩展、远程内容）隔离在最小权限边界内。
- 提示缓存友好：系统提示词前缀静态稳定，易变状态一律在消息尾部注入。

---

## 3. 总体架构

从外到内分七层。宿主层负责「我在哪个 Office 里」，Taskpane 应用层负责生命周期，Agent 运行时与工具层负责「推理 + 动作」，上下文/提示词层负责「让模型看得懂」，认证/连接层负责「能调用哪个模型」，存储层负责「一切要持久化的东西」。扩展、集成与本地桥接横跨其上，作为能力注入点。

```
┌────────────────────────────────────────────────────────────────┐
│  Office 宿主层（src/host/）                                     │
│  OfficeHost / WpsHost / BrowserHost · 全局检测 · onReady · 身份  │
├────────────────────────────────────────────────────────────────┤
│  Taskpane 应用层（src/taskpane/）                               │
│  boot · bootstrap · 会话管理 · tab 布局 · 上下文注入 · 状态栏      │
├────────────────────────────────────────────────────────────────┤
│  Agent 运行时（pi-agent-core / pi-ai）                          │
│  工具循环 · 流式输出 · 消息状态 · 模型抽象                       │
├────────────────────────────────────────────────────────────────┤
│  工具层（src/tools/）                                           │
│  注册表(registry) · 执行策略 · 协调器包装 · 连接预检 · 输出截断   │
├────────────────────────────────────────────────────────────────┤
│  上下文 / 提示词层（src/context/ src/prompt/）                  │
│  蓝图缓存 · 变更追踪 · system-prompt 构建 · 自动注入             │
├────────────────────────────────────────────────────────────────┤
│  认证 / 连接层（src/auth/ src/connections/ src/models/）        │
│  OAuth PKCE · API Key · 自定义网关 · CORS 代理 · 模型排序        │
├────────────────────────────────────────────────────────────────┤
│  存储层（src/storage/）—— IndexedDB「pi4office」6 个对象存储      │
└────────────────────────────────────────────────────────────────┘
      ↑ 扩展系统(extensions/) · 集成(integrations/) · 桥接(pkg/, scripts/)
```

### 3.1 三类运行边界

整个系统跨越三个执行边界，安全模型围绕它们展开（见[第 15 章](#15-安全模型)）：

| 边界 | 位置 | 说明 |
|---|---|---|
| 宿主边界 | Office/WPS 的 WebView 或普通浏览器 | CSP 收紧、无 unsafe-eval；runInWebView 兼容策略 |
| 本地进程边界 | 本机 Python 桥 / tmux 桥 / CORS 代理 / 安装器服务 | 仅 loopback 客户端、origin 白名单、目标主机白名单、负载与超时限制 |
| 远程边界 | 模型 API、Web Search、MCP 服务器、OAuth 端点 | CORS 代理仅对白名单目标放行，秘密最小化 |

### 3.2 依赖方向

- 上层通过显式接口调用下层；工具层是唯一被 Agent 运行时直接执行的边界。
- 注册表、协调器、恢复日志均为模块级单例，便于测试注入与替换。
- Node-only 依赖一律打浏览器 stub（vite 别名），保持 WebView 兼容（见 [16.3](#163-浏览器兼容策略)）。

---

## 4. 宿主抽象层（Office / WPS / Browser）

宿主层把「宿主是什么、何时就绪、工作簿身份是什么」收敛成一个接口 `SpreadsheetHost`（`src/host/types.ts`）。应用逻辑不直接触碰 Office.js 或 WPS 全局对象，全部经由该抽象。

### 4.1 宿主类型与检测

- 宿主类型 `SpreadsheetHostKind = office | wps | browser`；应用类型 `OfficeAppType = excel | word | unknown`。
- 同步检测 `detectSpreadsheetHost()`（`src/host/detection.ts`）：**WPS 优先**（WPS 不运行 Office.js，其全局为 `wps` 或 `Application`），其次 `Office` 全局，最后回退 `browser`。

### 4.2 启动解析与「永不降级」不变量

- `resolveSpreadsheetHostForBoot()`（`src/host/boot.ts`）：`office` 时创建 OfficeHost 并设 3 秒超时，等待 `Office.onReady`。
- **核心不变量**：若 Office 全局存在但 onReady 超时，仍以 `reason=office-timeout` 保留 OfficeHost，**绝不降级为 BrowserHost**——Office 启动慢时仍可延迟解析工作簿身份。
- 浏览器 / WPS 宿主同步就绪，无超时。初始化另有 60 秒硬截止，超时展示致命错误。

### 4.3 工作簿身份与隐私

- `getWorkbookContext()` 从宿主读取文档标识；Excel 由 `Office.context.document.url` 哈希派生，WPS 由 `Application.ActiveWorkbook.FullName` 派生（前缀 `wps_path_sha256:`）。
- **核心不变量**：绝不持久化或暴露原始文档 URL；身份用于隔离会话/恢复/工作簿关联，非用于追踪内容。

### 4.4 三个宿主实现

| 实现 | 行为要点 | 身份来源 |
|---|---|---|
| `OfficeHost` | `office.onReady` 驱动；主题解析；应用类型 excel/word | `document.url` 哈希 |
| `WpsHost` | 同步就绪；经 `wps/jsapi.ts` 访问 ET 对象模型；主题待支持 | `ActiveWorkbook.FullName` 哈希 |
| `BrowserHost` | 纯浏览器 / UI 画廊回退；应用类型 unknown | `unknown` 占位身份 |

- **WPS 全局访问只收敛在 `src/host/wps/jsapi.ts`** 一个文件（`window.wps` / `globalThis.Application`），便于测试时注入假的 WPS API 图。
- 会话-工作簿关联（`sessionStorage`）由 `settingsBackedSessionStorage` 适配到 settings 存储，三宿主共用。

---

## 5. 启动与引导流程

### 5.1 boot.ts —— UI 挂载前的兼容补丁

`taskpane.ts` 第一行强制导入 `./boot.js`，先后安装：

1. 加载第一方主题 CSS（`src/ui/theme.css`：tokens、preflight、组件样式）。
2. `installBedrockProviderStub()`：为 pi-ai 的 Node-only Bedrock provider 安装浏览器侧不支持 stub（lazy-load 兼容）。
3. `installCryptoRandomUUIDPatch()`：为缺失 `crypto.randomUUID` 的 WebView（如 WPS 12.1.0.26200）填充实现。
4. `installMarkedSafetyPatch()`：注入 marked 渲染钩子，阻断不安全链接协议（XSS 防线之一）。
5. `installThemeModeSync()`：把宿主 Office/WPS 深浅色主题同步到 pi-web-ui 的 `class="dark"`。

### 5.2 初始化时序

```
taskpane.ts
  ├─ import ./boot.js（最先，CSS + 补丁）
  ├─ import ./ui/register-components.js（注册第一方 web components）
  ├─ import ./ui/tool-renderers.js / message-renderers.js
  └─ bootstrapTaskpane()（src/taskpane/bootstrap.ts）
        ├─ 渲染 loading UI
        ├─ installProcessEnvShim() + installFetchInterceptor()（CORS 代理拦截）
        ├─ resolveSpreadsheetHostForBoot({ officeReadyTimeoutMs: 3000 })
        ├─ setCurrentSpreadsheetHost(host)   ← 全局单例
        └─ initTaskpane({ appEl, errorRoot })（src/taskpane/init.ts）
              ├─ initAppStorage() → IndexedDB「pi4office」
              ├─ 恢复会话标签页 / 布局
              ├─ buildRuntimeSystemPrompt()（动态系统提示词）
              ├─ createRuntime() → new Agent({ systemPrompt, model, tools,
              │                   transformContext, streamFn })
              └─ 注册命令、扩展、连接、技能
```

### 5.3 运行时与 Agent 装配

- 每个会话标签对应一个 `SessionRuntime`，内部持有一个 Agent 实例（pi-agent-core）。
- `transformContext = createContextInjector(changeTracker)`，在每轮消息发送前被框架回调（见[第 8 章](#8-上下文与提示词管理)）。
- `tools` 经三层包装管线装配（见 [6.4](#64-三层包装管线)）；模型由 `BrowserModelRuntime` 按当前 provider 凭据解析（见[第 9 章](#9-模型抽象与认证)）。

---

## 6. 工具系统

工具系统是「Agent 能对文档做什么」的唯一出口。它以注册表为单一事实来源，经执行策略分类、三层包装管线装配，最终在每轮请求中执行。

### 6.1 注册表作为单一事实来源

- 工具名的权威列表在 `src/tools/names.ts`：**17 个核心工具 = 14 个 Excel 专属 + 3 个共享**（`instructions` / `conventions` / `skills`）。
- `src/tools/registry.ts` 以 `CoreToolFactory` 字典把名字映射到 `create*Tool()` 工厂；参数类型在列表边界有意擦除，运行时由各工具自校验 TypeBox schema。
- **`src/tools/capabilities.ts` 为每个核心工具声明 `tier`、`category`（11 类：read/write/navigate/structure/format/inspect/view/collaboration/instructions/recovery/skills）与 `promptDescription`，并派生四条下游**：
  - 系统提示词的核心工具段落（`buildCoreToolPromptLines` → `system-prompt.ts` `buildToolsSection`）。
  - UI 渲染器与「人话化」参数（`TOOL_UI_METADATA` 标记 `renderer`/`humanizer`）。
  - 工具披露包 `TOOL_DISCLOSURE_BUNDLES`：按场景（core/analysis/formatting/structure/comments/full）按需增量暴露工具定义。
  - 主机适配 `host-selection.ts`：WPS 覆盖 `get_workbook_overview`/`read_range`/`write_cells` 的实现或插入占位工具。

### 6.2 执行策略：read / mutate，content / structure

`src/tools/execution-policy.ts` 将每个工具调用分类，决定加锁、检查点与蓝图失效行为：

- `ALWAYS_READ_TOOLS`（16 个）：读、`instructions`/`conventions`（改本地配置）、外部桥（`tmux`/`python_run`/`libreoffice_convert`/`web_search`/`fetch_page`/`mcp`/`files`/`extensions_manager`）。
- `ALWAYS_MUTATE_TOOLS`（10 个）：`write_cells`、`fill_formula`、`modify_structure`、`format_cells`、`conditional_format`、`python_transform_range`、`execute_office_js`、`execute_wps_js`。
- 双模工具按 `action` 参数判定：`view_settings(get)` / `comments(read)` / `charts(list|get_image)` / `workbook_history(restore 才 mutate)`。
- **未知工具默认 `mutate` —— 保守安全回退。**
- 上下文影响：仅结构性写操作（`modify_structure`、`execute_office_js`/`execute_wps_js`、`view_settings` 隐藏/显示工作表、`charts` 创建/删除）返回 `structure`，触发蓝图失效；其余写操作为 `content`。

### 6.3 恢复检查点

- 涉及写入的工具在自身 `execute` 内主动快照受影响单元格到恢复日志（`src/workbook/recovery-log.ts` 单例，持久化于 settings key `workbook.recovery-snapshots.v1`）。
- 快照种类 6 种：`range_values` / `format_cells_state` / `modify_structure_state` / `conditional_format_rules` / `comment_thread` / `chart_state`。
- 上限：`MAX_RECOVERY_CELLS=20,000`、`MAX_RECOVERY_ENTRIES=120`、保留区间 `[5,120]`。
- 回滚按种类分派；**核心不变量**：每次 restore 生成「反向快照」，支持二次回滚撤销恢复；恢复也经协调器 `runWrite` 串行化（见 [7.4](#74-写入后的副作用)）。
- 快照创建时派发 DOM 事件 `PI_WORKBOOK_SNAPSHOT_CREATED_EVENT`，供 UI 提示。

### 6.4 三层包装管线

`src/taskpane/init.ts` 的 `buildRuntimeCapabilities()` 按固定顺序装配：

```
裸工具列表（核心 + 集成 + 扩展）
  → 0. applyExperimentalToolGates()   （实验性工具门控，如 execute_office_js / python 桥审批）
  → 1. withWorkbookCoordinator()      （读直通 / 写串行 + Confirm 审批 + 写入后通知观察者）
  → 2. withConnectionPreflight()      （web_search/mcp 等连接前置检查，未配置时返回引导错误）
  → 3. applyToolOutputTruncation()    （按模型上下文窗口截断输出）
```

#### 6.4.1 withWorkbookCoordinator

- 读走 `coordinator.runRead` 直接执行；写先 `requireMutationApprovalIfNeeded()`（safe/Confirm 模式弹审批框，yolo/Auto 模式直通），再 `coordinator.runWrite` 串行执行。
- 写入提交后回调 `mutationObserver.onWriteCommitted(event)`，仅 `impact=structure` 时 `invalidateBlueprint(workbookId)`。

#### 6.4.2 withConnectionPreflight

- 对声明 `requiresConnection` 的工具，执行前检查 `ConnectionManager` 快照；非 connected 返回结构化 `connection_error`（含 `setupHint` 引导）。
- 执行抛错疑似认证失败时，脱敏错误消息并标记运行时认证失败，返回 `connection_auth_failed`。

#### 6.4.3 输出截断

- 默认窗口 2,000 行 / 50KB；按模型上下文窗口比例缩放（基准 128K，下限 200 行 / 8KB）。
- `head` 策略（保留开头）用于多数工具；`tail` 策略用于 `python_run`/`tmux`/`mcp`/`execute_office_js`/`execute_wps_js`（保留末尾）。
- 超过 512KB 的完整输出落盘到工作区 `.tool-output/<ts>-<tool>-<callid>.txt`，路径回写到截断详情供引用。

### 6.5 工具结果双通道

- `result.content` 是人类可读的 markdown（供渲染），`result.details` 是稳定的机器元数据（供 UI/策略消费）。
- 该约定写入 AGENTS.md 的协作规范：新增工具时必须同步更新 UI 渲染器、人话化参数、工具披露、系统提示词五处（见[第 19 章](#19-开发与测试)）。

---

## 7. 并发与一致性：工作簿协调器

> **核心不变量**：并行推理，串行化工作簿写入（parallel reasoning, serialized workbook mutation）。

### 7.1 为什么需要

一个工作簿可开多个会话标签，每标签一个 Agent 实例并发推理。若它们同时写单元格，会出现竞争与相互覆盖。协调器保证同一工作簿的写操作严格按入队顺序执行，而读操作不受阻塞。

### 7.2 队列机制

- 每工作簿一个 `WorkbookQueueState`：`revision`（每次写成功 +1）、`running` 互斥位、`activeWrite`、待执行队列。
- `runWrite` 把操作封装为 Promise 入队，`processQueue` 仅当 `running=false` 时取出队首执行，`finally` 复位并递归处理下一个。
- `runRead` 不排队直接执行，同时发出 started/completed/failed 事件供 UI 追踪。
- `revision` 同时充当「写入代际」，上下文注入据此判断蓝图是否过期。

### 7.3 Confirm 审批

- 执行模式由 `getExecutionMode` 决定，默认 yolo（Auto）；safe（Confirm）模式在每次工作簿变更前弹审批框。
- 审批消息含模式标签、工具名、Range 与提示；请求经 `requestRuntimeToolApproval → requestConfirmationDialog` 联动 UI。

### 7.4 写入后的副作用

- 写提交 → `mutationObserver.onWriteCommitted`：仅 `structure` 影响失效蓝图（触发下轮重新注入蓝图）。
- 恢复/回滚也经 `runWrite` 串行化，避免与其它写交错。
- 协调器事件订阅驱动 UI 锁状态：`queued→waiting_for_lock`、`started→holding_lock`、`completed/failed→idle`。

---

## 8. 上下文与提示词管理

### 8.1 自动上下文注入

`createContextInjector(changeTracker)` 返回 `transformContext`，在每轮用户消息发送前被 Agent 框架调用。它把一条 `[Auto-context]` 前缀的 user 消息插到最近用户消息之前，按序注入四类内容：

1. **工作簿蓝图**：结构刷新时注入（见 [8.2](#82-蓝图缓存与刷新决策)），带 `[Workbook context refresh: <reason>]` 标记。
2. **工作区文件摘要**：签名变化时注入，来自 `workspace.getContextSummary()`。
3. **选区上下文**：当前选区及上下各 5 行的 markdown 表格。
4. **改动摘要**：`changeTracker.flush()` 返回自上次消息以来的用户编辑，去重分组、上限 50 条。

### 8.2 蓝图缓存与刷新决策

- `src/context/blueprint.ts` 以 `Map` 缓存蓝图文本与单调递增修订号；`invalidateBlueprint` 删缓存并自增修订号。
- `decideWorkbookContextRefresh()` 综合 `workbookId` 与修订号决定刷新原因：`workbook_switched` / `blueprint_invalidated` / `context_missing` / `initial`。
- 刷新决策的闭包状态 `lastInjectedWorkbookId` / `lastInjectedBlueprintRevision` 保证「仅在结构变化时重读工作簿」，避免每轮全量扫描。

### 8.3 系统提示词构建

`src/prompt/system-prompt.ts` 的 `buildSystemPrompt()` 按序拼装 13 段，且按宿主分支（Excel/Word）差异化：

```
① 身份(Identity)        ⑥ 本地服务(Local Services)    ⑪ 约定(Conventions, Excel)
② 规则(Rules)           ⑦ 可用技能(Skills)            ⑫ 自定义格式预设(Custom Presets, Excel)
③ 执行模式(Execution)   ⑧ 工具(Tools)                 ⑬ 约定覆盖(Convention Overrides)
④ 活跃集成(Integrations)⑨ 工作区(Workspace)
⑤ 连接状态(Connections) ⑩ 工作流(Workflow)
```

- 工具段落：Excel 用 capabilities.ts 动态生成的 `CORE_TOOL_PROMPT_LINES`；Word 用硬编码工具清单。
- 工作流段落：Excel 9 条、Word 7 条规则，均强调「先读后写」「最小改动范围」。
- 连接段落按 Connected / Missing / Needs attention 三组列出；本地服务段落含 python/tmux 桥健康检查结果（固定排序）。

### 8.4 提示缓存友好设计

- 系统提示词前缀保持静态稳定；时间戳、随机 ID 等易变信息一律走消息尾部注入，避免缓存键漂移。
- 工具列表顺序确定；用 fingerprint + extension tool revision 语义判断缓存何时必须失效（`src/auth/prefix-churn.ts`）。
- `PrefixChangeTracker` 对每会话的模型/提示/工具指纹做 FNV 哈希，变更时通知运行时刷新前缀。

---

## 9. 模型抽象与认证

### 9.1 浏览器内模型运行时

- `BrowserModelRuntime` 用 `ProviderCredentialsStore` + `BrowserModelCatalogsStore` 调 pi-ai 的 `createModels()`，以浏览器原生方式注册内置 provider。
- 缺 `apiKey` 认证的 provider 包一层 `createBrowserAdapterProvider()`；自定义网关与扩展 provider 动态注册并做 ID 冲突检测。
- 模型目录缓存按当前 api/baseUrl 重绑，读取时校验清理模型元数据。

### 9.2 模型选择与排序

- `compareModels` 排序：provider 优先级 → 家族 → 版本新旧 → 字母序；版本解析可处理 slashes/dashes/日期后缀。
- 模型选择器 `orderModelsForSelector`：当前模型置顶，其次各 provider「推荐」最新模型，其余确定性排序。
- 默认模型 `pickDefaultModel` 优先级：openai-codex → anthropic(Opus>Sonnet>Fable) → Google → 自定义 → 任意可用 → 回退 `gpt-5.6-sol`。
- `active-providers` 集合按「用户持有凭据的 provider」过滤不可运行的模型。

### 9.3 连接管理

- `ConnectionManager` 以 `ownerId` 组织 provider 定义注册表；状态机 `missing → connected → invalid → error`。
- **秘密脱敏**：错误消息中仅 ≥4 字符、全匹配、转义正则替换为 `….`，且绝不向 UI 暴露原始秘密（`getSecretFieldPresence` 返回布尔掩码）。
- 连接定义声明 `authKind`（api_key/bearer_token/oauth/custom）与 `httpAuth`（header 模板 + 允许主机）。

### 9.4 浏览器 OAuth（PKCE）

- 纯浏览器 PKCE：随机 verifier + SHA-256（Web Crypto subtle 优先，WPS/dev HTTP 回退到旁加载的纯 JS 实现）。
- 五个 provider：`anthropic`、`openai-codex`、`google-gemini-cli`、`google-antigravity`、`github-copilot`；凭证存 IndexedDB（键 `oauth.<providerId>`）。
- 本地回调端口：anthropic:53692、openai-codex:1455、gemini:8085、antigravity:51121；OAuth 捕获由 CORS 代理的 `/oauth/callback/<provider>` 中转。
- **核心不变量**：OAuth 刷新仅在任务窗格执行；`ProviderCredentialsStore` 拒绝在 `modify()` 中刷新（浏览器 OAuth 必须由 `restore.ts` 刷新循环处理）。

### 9.5 API Key 与自定义网关

- API Key 存 `ProviderKeysStore`（IndexedDB，纯文本键值）；开发态可经 `/__pi-auth` 复用 pi 的 `~/.pi/agent/auth.json`（仅 loopback + localhost Host 头放行）。
- 自定义 OpenAI 兼容网关以 `pi-openai-gateway:` 前缀注册为独立 provider，默认上下文窗口 16,384 令牌；支持 Ollama/llama.cpp/vllm/lmstudio 自动发现。
- `openai-codex` 有意路由到 `chatgpt.com/backend-api` 而非 `api.openai.com`（`provider-map.ts` 注释明确）。

### 9.6 CORS 代理

- `installFetchInterceptor()` 猴子补丁 `window.fetch`：开发态把外部 URL 重写为 Vite 反向代理路径（`/oauth-proxy/*`、`/api-proxy/*`）；生产态仅对 OAuth/令牌端点按 `/?url=<target>` 模式走用户配置的代理。
- 代理校验：`validateOfficeProxyUrl` 拒绝 WebView HTTPS→HTTP 混合内容；`probeProxyReachability` 与 codex WebSocket 桥探针做健康检查。
- `dev-auth-policy.ts` 保护 `/__pi-auth` 端点：仅环回地址 + localhost Host 头。

### 9.7 运行时模型对账

- provider 配置变更后，若当前运行时模型指向的 provider 失去凭据，`resolveRuntimeModelSwap` 在运行时空闲时自动切到可用默认模型（问题 #553）。

---

## 10. 存储层

### 10.1 单一 IndexedDB 数据库

- `initAppStorage(dbName=pi4office)` 创建版本 2 的单一数据库，内含 6 个对象存储：`settings`、`provider-keys`、`sessions`、`sessions-metadata`、`custom-providers`、`model-catalogs`。
- 共享 `IndexedDBStorageBackend` 支持跨存储原子事务、`keys()` 前缀过滤、配额查询与 `navigator.storage.persist()`。
- `AppStorage` 是统一门面，全局单例 `getAppStorage()`/`setAppStorage()` 便于测试注入。

### 10.2 各存储职责

| 存储 | 内容 | 说明 |
|---|---|---|
| `SettingsStore` | 应用设置：主题、代理、OAuth、特性开关、tab 布局、恢复日志 | 通用键值 |
| `ProviderKeysStore` | LLM API Key（按 provider 名） | 纯文本；`ProviderCredentialsStore` 适配 pi-ai 的 `CredentialStore` 接口 |
| `SessionsStore` | 完整会话消息 + 轻量元数据 | 事务性双写；按 `lastModified` 索引排序 |
| `CustomProvidersStore` | 自定义网关定义 | 含自动发现 / 手动两种类型 |
| `ModelCatalogsStore` | 模型发现缓存 | 实现 pi-ai `ModelsStore` 接口，读取时校验 |
| （auth） | OAuth 凭证 `oauth.<providerId>` | 见 [9.4](#94-浏览器-oauthpkce) |

### 10.3 存储卫生与边界

> **注意**：IndexedDB 对同源脚本可读，是**存储卫生而非 XSS 边界**。真正防线是 CSP + marked 安全补丁 + 沙箱（见[第 15 章](#15-安全模型)）。

---

## 11. 会话管理

### 11.1 多会话标签

- `SessionRuntimeManager` 以 `runtimes` Map + `runtimeOrder` 管理多个 `SessionRuntime`；每个运行时持有一个 Agent 实例与独立队列显示。
- 标签号从 1 起单调分配；切换/关闭/重排都有对应方法，关闭活动标签自动切到邻近标签，最后一个标签不可关闭。

### 11.2 自动保存 / 恢复

- **核心不变量**：仅在生成第一条助手消息后才持久化会话，避免空会话堆积。
- `saveSession` 聚合 token/成本统计、构建 2048 字符预览、事务性保存 SessionData+Metadata，并更新工作簿-会话关联。
- `restoreLatestSession` 优先按工作簿 ID 恢复关联会话，否则回退全局最新；跳过空会话。

### 11.3 布局与最近关闭

- tab 布局持久化于 settings 键 `workbook.tabLayout.v1.<workbookId>`（无工作簿时用 `__global__`）。
- `RecentlyClosedStack` 是 10 项 LIFO，支撑 `/reopen` 恢复最近关闭的标签。
- `/resume` 以新标签恢复，`/resume-here` 替换当前标签；`/name` 设置显式标题。

---

## 12. 斜杠命令系统

### 12.1 命令注册表

- `SlashCommand` 声明 name/description/source（`builtin`|`extension`|`integration`|`prompt`）、execute(args)、可选 `busyAllowed` 与 `enabled()`。
- `CommandRegistry` 为单例，支持 register/unregister/get/list（对名称与描述做大小写不敏感子串过滤）。

### 12.2 内置命令分组

按菜单顺序注册：模型 `/model`、设置 `/settings` `/yolo` `/proxy`、插件、工具、技能、文件、实验 `/experimental`、调试、剪贴板、导出、会话身份 `/name` `/share-session`、帮助、扩展、会话生命周期 `/new` `/resume` `/resume-here` `/history` `/backup` `/reopen` `/revert`、压缩 `/compact`。

### 12.3 忙碌命令策略

- 运行忙碌（流式中）时仅放行 `BUSY_ALLOWED_COMMANDS` 白名单（`compact`/`new`/`rules`/`resume`/`history`/`reopen`/`yolo`/`extensions`/`plugins`/`skills`/`files`/`tools`）；扩展命令默认 `busyAllowed`。
- `executeSlashCommand` 返回 `not-found` / `busy-blocked` / `missing-queue` / `queued` / `executed`；`/compact` 支持排队执行。

### 12.4 实验性功能门控

| 内部 id / 命令 | 作用 | 默认 |
|---|---|---|
| `ui_dark_mode` | 界面深色模式 | 关闭 |
| `remote_extension_urls` | 允许加载远程 URL 扩展 | 关闭 |
| `extension_permission_gates` | 执行扩展能力权限 | 关闭 |
| `extension_sandbox_runtime` | 沙箱回滚开关（开启=回退宿主运行时） | 关闭 |
| `extension_widget_v2` | Widget API v2 | 关闭 |

- 特性标志持久化到 **localStorage（而非 IndexedDB）**，保证在存储初始化前可用；变更派发 `pi:experimental-feature-changed` 事件。

---

## 13. 扩展系统

### 13.1 注册表与信任分级

- 扩展注册表持久化于 settings 键 `extensions.registry.v2`（v1 自动迁移）。来源分 module specifier（本地/远程 URL）与 inline 代码。
- 四个信任级：`builtin` / `local-module` / `inline-code` / `remote-url`。内置默认扩展 `builtin.snake`（默认启用）。

### 13.2 权限模型

- 20 个能力键：`commandsRegister`、`toolsRegister`、`modelsRegister`、`agentRead`、`agentEventsRead`、`uiOverlay`、`uiWidget`、`uiToast`、`llmComplete`、`httpFetch`、`storageReadWrite`、`connectionsReadWrite`、`connectionsSecretsRead`、`clipboardWrite`、`agentContextWrite`、`agentSteer`、`agentFollowUp`、`skillsRead`、`skillsWrite`、`downloadFile`。
- trusted（builtin/local-module）默认全开，但 `connectionsSecretsRead`/`agentContextWrite`/`agentSteer`/`agentFollowUp`/`skillsWrite` 默认关闭；untrusted（inline/remote）默认仅开少量 UI/存储/剪贴板能力。
- 权限执行由 `/experimental on extension-permissions` 门控（默认关闭）。

### 13.3 运行时模式与沙箱

- `resolveExtensionRuntimeMode`：builtin/local-module → host；inline/remote → 默认 `sandbox-iframe`，可经 `/experimental on extension-sandbox-rollback` 回退 host。
- 沙箱以 `srcdoc` iframe（`sandbox="allow-scripts"`，隐藏）运行不可信代码；通信经 `MessageChannel`，首次 postMessage 用通配 origin，之后全部走专用 `MessagePort`；RPC 超时 15s。
- 每个 RPC 方法执行前 `assertCapability()` 门控；`api.agent.raw` 在沙箱不可用（有意收窄 API 面）。
- UI 树投影深度限制 12 层、标签白名单（div/span/p/strong/em/code/pre/ul/ol/li/h1–h6/button）、`data-pi-action` 清洗为正则 `[A-Za-z0-9:_-]{1,48}`。
- 沙箱暴露方法覆盖命令/工具/模型注册、`llm_complete`、`http_fetch`、存储、剪贴板、agent 注入/steer/follow_up、技能、连接、toast/overlay/widget、订阅 agent 事件。

### 13.4 安全边界

- 远程 URL 默认禁止导入（`extension-source-policy.ts` 分类 `local-module` / `blob-url` / `remote-url` / `unsupported`），需 `/experimental on remote-extension-urls`。
- 扩展 `http_fetch` 阻断 localhost/`.localhost`/`.local`/`0.0.0.0`/`127.0.0.1`/`::1` 及私有网段（`runtime-manager-helpers.ts`）。
- `RuntimeManager` 失败隔离：每个扩展独立捕获错误记录到 `lastErrors`；工具名冲突检测（保留名 + 跨扩展唯一）；`extensionToolRevision` 递增作为提示缓存指纹。

### 13.5 extensions_manager 工具与 Widget API

- `extensions_manager` 支持 `list` / `install_code` / `set_enabled` / `reload` / `uninstall`；`install_code` 默认替换同名扩展。
- Widget API v2（`upsert`/`remove`/`clear`）由 `/experimental on extension-widget-v2` 门控；placement 分 `above-input` / `below-input`，min/max 高度由宿主钳制在 72–640px，按 ownerId 分组。

---

## 14. 集成与本地桥接

### 14.1 集成目录

- 内置两个集成：`web_search`（工具 `web_search` + `fetch_page`，默认启用，provider：Jina 默认 / Firecrawl / Serper / Tavily / Brave）与 `mcp_tools`（工具 `mcp`，Alpha，仅支持 HTTP MCP 服务器，默认不启用）。
- 集成状态按 session 与 workbook 两级作用域持久化；workbook 未配置时继承目录默认；全局门控 `external.tools.enabled` 可一键关闭所有外部集成。
- 每个集成定义注入 system prompt 的 `instructions`、`agentSkillName` 与 `warning`。

### 14.2 Web Search 与 fetch_page

- `web_search` 对配置的 keyed provider 失败时自动回退到 Jina 并给出 ⚠️ 提示（前提存在 Jina key）。
- `fetch_page` 默认 `max_chars=12,000`（1k–50k），同域名节流 1s。

### 14.3 Python / LibreOffice 桥

- 默认路径是浏览器内 **Pyodide**（Web Worker + WASM，CDN 版本 v0.27.7，懒加载，超时 30s，崩溃重建）。
- native 路径是 `python-bridge-server.mjs`（端口 3340，HTTPS，可带 bearer token）：端点 `GET /health`、`POST /v1/python-run`、`POST /v1/libreoffice-convert`。
- 防护：仅 loopback、body ≤512KB、代码 ≤40K、输出 ≤256KB、Python 超时 10–120s、LibreOffice 1–300s；Python 以 `-I` 隔离模式 exec 受限 scope，结果经 `__PI_FOR_EXCEL_RESULT_JSON_V1__` 标记提取。
- LibreOffice 以 argv 数组调用 `soffice --headless --convert-to <fmt> --outdir <tmp> <input>`（无 shell 插值），自动查找产物并复制到输出路径；支持 csv/pdf/xlsx。
- `libreoffice_convert` 分类 read/无工作簿影响；`python_transform_range` 分类 mutate/content（写回工作簿）。打包为 `npx pi4office-python-bridge`。

### 14.4 tmux 桥

- `tmux-bridge-server.mjs`（端口 3341，stub/real 两种模式）：`GET /health`、`POST /v1/tmux`；动作 list/create/send_keys/capture_pane/send_and_capture/kill_session。
- 会话名正则 `^[A-Za-z0-9][A-Za-z0-9._:-]{0,63}$`；`cwd` 必须绝对且存在；固定 socket 路径，目录权限 `0o700`。
- `tmux` 工具分类 read/无工作簿影响（不取写锁）。打包为 `npx pi4office-tmux-bridge`。

### 14.5 CORS 代理服务器

- `cors-proxy-server.mjs`（默认 3003，格式 `/<path>?url=<encoded>`）：仅 loopback 客户端 + origin 白名单；目标主机两级白名单检查（见 [15.3](#153-代理目标策略)）；剥离 hop-by-hop 头、不转发 cookie、不转发 Content-Encoding/Length。
- 支持 GitHub Enterprise OAuth/Copilot 特殊路径、Codex WebSocket 桥（`pi_transport=codex-websocket` → SSE）、每 provider 独立 localhost 回调捕获服务器（可经 `OAUTH_CALLBACK_SERVER=0` 关闭）。
- 危险开关 `ALLOW_ALL_TARGET_HOSTS=1` 不可提交仓库；`ALLOWED_CLIENT_CIDRS` 拒绝 `/0`（fail-closed）。

---

## 15. 安全模型

### 15.1 威胁模型概要

`docs/security-threat-model.md`（v1）定义四类敏感数据与四个信任边界：任务窗格 WebView、本地辅助服务、远程 provider、扩展运行时。四类主要威胁与控制：

| 威胁 | 攻击面 | 控制 |
|---|---|---|
| XSS / 内容注入 | 模型输出 markdown、扩展 UI 树 | marked 安全补丁阻断危险协议、图片渲染为链接、CSP、UI 标签白名单 |
| 令牌泄露 | IndexedDB 凭据 | 仅 IndexedDB（无 localStorage 回退）、断连清理、连接秘密脱敏 |
| 本地代理滥用 | CORS 代理 / 桥 | loopback 客户端、origin 白名单、目标主机白名单、负载与超时限制 |
| 扩展代码执行 | 扩展运行时 | 远程 URL 默认阻断、untrusted 默认沙箱、权限门控、回滚开关 |

### 15.2 CSP

- `vercel.json` 为 `/src/taskpane.html` 定义严格 CSP：`default-src 'none'`；`script-src 'self' blob:` + Office 与 Pyodide CDN；`object-src`/`base-uri`/`form-action` 全部 `'none'`。
- `connect-src` 仅白名单模型 API、搜索 provider、github、`localhost:*` 等；测试 `taskpane-csp.test.mjs` 守护此配置。
- 因 CSP 无 `unsafe-eval`，Ajv 的 `new Function()` 被浏览器 stub 掉（`src/stubs/ajv.ts`），触发 pi-ai 的跳过校验回退。

### 15.3 代理目标策略

- `evaluateTargetHostPolicy()`（`scripts/proxy-target-policy.mjs`）两级检查：hostname 白名单 → DNS 解析 IP 是否为 loopback/私有。
- **核心不变量**：显式配置 `ALLOWED_TARGET_HOSTS` 时 `requireAllowlistForOverriddenTargets=true`，防止 `ALLOW_LOOPBACK`/`ALLOW_PRIVATE` 绕过白名单。
- 客户端策略 fail-closed：非法 CIDR 条目直接致命错误。

### 15.4 证书与清单权限

- 开发 HTTPS 用仓库根 `key.pem`/`cert.pem`（`mkcert localhost` 生成）；集中部署支持 `TLS_KEY_PATH`/`TLS_CERT_PATH`。
- 加载项清单声明 `Permissions=ReadWriteDocument`，Hosts=Workbook+Document，VersionOverrides 含两者 DesktopFormFactor 的 ribbon 按钮。

### 15.5 已知限制

- IndexedDB 非 XSS 边界（同源脚本可读）。
- builtin/local-module 扩展以 host 运行时信任运行。
- 能力权限执行当前 opt-in（`extension-permissions` 实验开关）。
- 沙箱有意不提供 `api.agent.raw`，收窄但不为零的 API 面。

---

## 16. 构建与发布

### 16.1 开发服务器

- `vite.config.ts`：dev server 固定 `https://localhost:3141`（strictPort，IPv4/IPv6 双栈绑定 `::`），因为清单硬编码该地址。
- 反向代理：`/oauth-proxy/*`（Anthropic/Google/GitHub OAuth）+ `/api-proxy/*`（Anthropic/OpenAI/Google/Cloud Code），剥离浏览器头避免 CORS 触发，并修复 Cloud Code 的 `%3A`→`:` 编码。
- `piAuthPlugin` 开发态暴露 `/__pi-auth`（仅 loopback），复用 `~/.pi/agent/auth.json` 免重登录。
- portless 模式：设 `DEV_HOST` 或 `PORTLESS_URL` 时 TLS 终结于代理，Vite 仅监听 loopback 纯 HTTP。

### 16.2 清单生成与验证

- `generate-manifest.mjs` 把 dev 清单所有 `https://localhost:3141` 替换为 `ADDIN_BASE_URL`，输出 `manifest.prod.xml`（根 + `public/`）。
- `validate-manifest.mjs`（`npm run validate`）校验清单；CI 强制执行。

### 16.3 浏览器兼容策略

- vite 别名把 Node-only 导入 stub 掉：`stream`（Anthropic SDK 传递依赖）、`ajv`/`ajv-formats`（CSP 无 unsafe-eval）。
- `resolve.dedupe=[marked]` 强制单一 marked 实例，保证安全补丁拦截所有解析。
- `build.external` 仅外部化 `/^node:/`，避免把常规依赖误判为 bare import 导致加载项启动失败。

### 16.4 部署与 CI

- Vercel：build 后 output `dist`；`/src/taskpane.html` 与 `manifest.prod.xml` 强制 no-store；附带安全响应头；`vercel-ignore-command` 决定 main/PR/手动构建。
- GitHub Pages 镜像工作流（`deploy-pages.yml`，`--base=/pi4office/`）用于 `*.vercel.app` 被阻断的网络（如中国大陆）。
- CI（`ci.yml`）：`npm ci` → `audit` → `check` → `test` → `build` → `validate`；另有 dependabot 自动合并、dependency-review。

### 16.5 Windows 安装器与 CLI 包

- `build-exe-installer.mjs`：构建前端 → 下载 Node portable + mkcert → 复制 dist/server/install 脚本 → 打包；`--exe` 用 NSIS `makensis` 编译安装器。
- 可发布 CLI 包：`npx pi4office-proxy`、`npx pi4office-python-bridge`、`npx pi4office-tmux-bridge`（`pkg/` 目录 + 同步脚本）。

---

## 17. WPS 支持（实验性）

### 17.1 现状

- 已实现 WPS Spreadsheets 后端，目标 `wps.cn` 个人版与 WPS 365 企业版；国际版 `wps.com` 暂不支持。
- 工具名 / label / description / 参数 schema 跨宿主保持一致；不支持的 WPS 工具以 `UnsupportedHostToolError` 快速失败。

### 17.2 工具矩阵

| 类别 | 工具 | 说明 |
|---|---|---|
| 支持 | `get_workbook_overview` / `read_range` / `write_cells` | 含覆盖保护与写回校验；无 WPS 自动备份 |
| 支持 | `instructions` / `conventions` / `skills` / `execute_wps_js` | 共享工具 + WPS 专属逃生舱 |
| 快速失败 | `workbook_history`、`fill_formula`、`search_workbook`、`modify_structure`、`format_cells`、`conditional_format`、`charts`、`trace_dependencies`、`explain_formula`、`view_settings`、`comments`、`execute_office_js`、`python_transform_range` | 阶段未实现，抛 `UnsupportedHostToolError` |
| 不变 | `python_run` / `tmux` / `libreoffice_convert` / `files` / `extensions_manager` 及集成 | 与宿主无关 |

### 17.3 打包与已知阻断

- `wps/` 目录：`index.html` 入口、`ribbon.xml`（无 XML declaration，规避发布验证器）、`main.js`、`jsplugins.xml.template`。
- 部署通道：`wpsjs publish`（推荐）或 `jsplugins.xml` + `oem.ini`（企业/OEM）。
- 已知产品验证阻断：WPS 个人版 12.1.0.26200（win-arm64ec）生成 `authaddin.json` `enable:false` + `jsaddinblockhost.ini`；x86 / WPS 365 可正常加载。

---

## 18. 目录结构与模块职责

| 目录 | 职责 | 关键符号 |
|---|---|---|
| `src/host/` | 宿主抽象与检测 | `detectSpreadsheetHost`、`SpreadsheetHost`、`OfficeHost`/`WpsHost`/`BrowserHost` |
| `src/taskpane/` | 应用装配、会话、上下文注入、状态栏 | `bootstrapTaskpane`、`init.ts`、`SessionRuntimeManager`、`createContextInjector` |
| `src/tools/` | 工具注册表、执行策略、包装、检查点 | `registry.ts`、`execution-policy.ts`、`with-workbook-coordinator.ts`、`recovery-log.ts` |
| `src/workbook/` | 工作簿身份、协调器、恢复 | `coordinator.ts`、`recovery/`、`session-association.ts` |
| `src/context/` | 蓝图缓存、变更追踪、选区 | `blueprint.ts`、`change-tracker.ts`、`selection.ts`、`window-budgets.ts` |
| `src/prompt/` | 系统提示词构建 | `system-prompt.ts` |
| `src/auth/` | OAuth/API Key/代理/网关/前缀追踪 | `pkce.ts`、`cors-proxy.ts`、`oauth-provider-registry.ts`、`custom-gateways.ts` |
| `src/connections/` | 集成连接状态机与秘密 | `manager.ts`、`store.ts` |
| `src/models/` | 模型排序、默认选择、运行时对账 | `model-ordering.ts`、`browser-model-runtime.ts` |
| `src/storage/` | IndexedDB 后端与各存储 | `init-app-storage.ts`、`local/indexeddb-storage-backend.ts` |
| `src/commands/` | 斜杠命令注册与执行 | `types.ts`、`builtins/`、`slash-command-execution.ts` |
| `src/extensions/` | 扩展注册、权限、沙箱、运行时 | `runtime-manager.ts`、`sandbox-runtime.ts`、`permissions.ts` |
| `src/integrations/` | Web Search / MCP 目录与开关 | `catalog.ts`、`store.ts` |
| `src/python/` | Pyodide 运行时（Worker） | `pyodide-runtime.ts`、`pyodide-worker.ts` |
| `src/compat/` | 浏览器兼容补丁 | `marked-safety.ts`、`crypto-random-uuid.ts`、`bedrock-provider-stub.ts` |
| `src/experiments/` | 实验性特性开关 | `flags.ts` |
| `src/ui/` | 侧边栏组件、工具/消息渲染、主题 | `register-components.ts`、`tool-renderers.ts`、`theme/` |
| `scripts/` | dev 脚本、桥服务器、check/构建脚本 | `python-bridge-server.mjs`、`cors-proxy-server.mjs`、`generate-manifest.mjs` |
| `pkg/` | 可发布 CLI 包与安装器 | `proxy/`、`python-bridge/`、`tmux-bridge/`、`installer/` |
| `tests/` | 单元 + 安全测试（约 50 文件） | `test:models` / `test:context` / `test:security` / `test:manifest` |
| `wps/` | WPS 加载项入口与 ribbon | `index.html`、`ribbon.xml`、`main.js` |

---

## 19. 开发与测试

### 19.1 常用命令

| 命令 | 作用 |
|---|---|
| `npm run dev` / `use` | HTTPS dev server（:3141）/ 起服务并侧载 |
| `npm run build` | 生产构建 → `dist/` |
| `npm run check` | lint + typecheck + 主题/CSS/一致性检查 |
| `npm run test` | test:models + test:context + test:security + test:manifest |
| `npm run validate` | 校验加载项清单 |
| `npm run proxy:https` | CORS 代理（OAuth 用） |
| `npm run build:installer` | 构建 Windows 安装器 |

### 19.2 测试套件

- `test:models`：模型排序、运行时模型对账、浏览器模型运行时。
- `test:context`：工具、上下文注入、会话、扩展、集成、恢复、协调器、输出截断、实验门控。
- `test:security`：代理目标/客户端策略、CORS、沙箱、OAuth、CSP、桥服务器安全。
- `test:manifest`：清单生成。

### 19.3 协作与质量规范（AGENTS.md / coding-standards.md）

- 新增/修改工具必须同步五处：`registry.ts`、`tool-renderers.ts`、`humanize-params.ts`、`tool-disclosure.ts`、`system-prompt.ts`。
- 用户可见 UI 字符串走 `t()`（i18n），`en.json` 为唯一事实；`t()` 绝不在模块作用域调用；agent 面向字符串不进 i18n。
- 禁止 `// @ts-ignore`、显式 `any`/`as any`；避免 Node-only import 与 barrel import。
- 提示缓存：工具列表顺序必须确定；系统提示词不用时间戳/随机 ID；用 fingerprint + 扩展工具修订号。
- pre-commit 钩子运行 `npm run check`；`--no-verify` 可绕过。
- 任何「完成/修复/通过」结论须附验证命令或结果摘要（见项目开发规范）。

---

## 20. 核心不变量速查与参考

### 20.1 核心不变量清单

- **并行推理、串行写入**：同一工作簿的写操作经协调器严格串行。
- **宿主永不降级**：Office 全局存在即保留 OfficeHost，即使 onReady 超时。
- **工作簿身份隐私**：绝不持久化原始文档 URL。
- **恢复可逆**：每次 restore 生成反向快照，支持二次回滚。
- **蓝图按需刷新**：仅 `structure` 级写操作使蓝图失效。
- **OAuth 刷新只在任务窗格**：`ProviderCredentialsStore` 拒绝在 `modify()` 内刷新。
- **提示缓存稳定**：系统提示词前缀静态，易变状态走消息尾部。
- **代理目标 fail-closed**：目标主机白名单不可被 loopback/私有豁免绕过。
- **会话惰性持久化**：仅生成首条助手消息后保存，避免空会话堆积。

### 20.2 参考文档

| 文档 | 内容 |
|---|---|
| `docs/security-threat-model.md` | 威胁模型 v1：四信任边界与四威胁 |
| `docs/python-bridge-contract.md` | Python/LibreOffice 桥契约（门控、Pyodide 回退、分类） |
| `docs/tmux-bridge-contract.md` | tmux 桥契约（动作、安全、分类） |
| `docs/central-proxy.md` | 组织托管集中 CORS 代理部署 |
| `docs/context-management-policy.md` | 上下文与提示缓存策略 |
| `docs/coding-standards.md` | 代理协作编码规范 |
| `docs/upstream-divergences.md` | 与上游 pi-mono 的差异 |
| `docs/wps-support.md` | WPS 支持计划与工具矩阵 |
| `docs/extensions.md` | 扩展开发指南（MVP） |
| `docs/deploy-vercel.md` | Vercel 托管部署步骤 |
| `src/tools/DECISIONS.md` | 工具行为决策记录 |
| `src/ui/README.md` | UI 架构说明 |

---

> **结语**：本文以 v0.10.0-pre 为基线。工具数量、注册表派生链、安全门控均可能随版本演进，阅读实现时以 `src/` 下注释与 AGENTS.md 为最新事实来源。

---

## 许可证

[MIT](../LICENSE) © Thomas Mustier
