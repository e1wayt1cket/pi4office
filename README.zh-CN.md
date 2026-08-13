# Pi for Office

[English](./README.md) | 简体中文

> [!NOTE]
> 本简体中文指南由 AI 翻译/生成，可能存在译文问题；如与英文文档不一致，请以英文版为准。欢迎提 [Issue](https://github.com/tmustier/pi4office/issues) 指正。

> 本文档是简要中文指南，涵盖 **Microsoft Office 加载项**的**安装**与**模型配置**，支持 Excel 和 Word。完整功能说明、开发者文档等请参阅[英文版 README](./README.md) 与 [docs/](./docs/README.md) 目录。WPS 表格支持另见英文文档 [docs/wps-support.md](./docs/wps-support.md)。

Pi for Office 是一款开源、多模型的 Microsoft Office AI 侧边栏加载项，由 [Pi](https://pi.dev) 驱动，支持 **Excel** 和 **Word**。

它是一个运行在 Office 内部的 AI 智能体：能读取你的文档、修改内容、进行联网研究——模型由你选择。既支持 Anthropic、OpenAI、Google Gemini、GitHub Copilot 的 API Key 或 OAuth 登录，也支持任何 **OpenAI 兼容接口**(如 DeepSeek、智谱 GLM、Ollama 本地模型等)。

**功能一览**(详见[英文版 README](./README.md#features)):

- **Excel**: 16 个内置电子表格工具——读写单元格、填充公式、全簿搜索、结构调整、单元格格式与条件格式、公式解释与依赖追踪、批注、自动备份等
- **Word**: 文档结构概览、读写文档内容、插入文本、搜索文档、格式化文本、添加批注
- 多模型支持，对话中可随时切换模型
- 会话管理：每个文档多个会话标签页、自动保存/恢复、历史记录
- 自动上下文注入：AI 自动获知文档结构和当前选区，无需手动描述
- 每次修改前自动创建检查点，出错可一键回滚
- 斜杠命令、扩展系统、联网搜索 + MCP 集成

---

## 安装

### 前置准备

#### 1. 确认系统要求

在开始安装前，请确保您的环境满足以下要求：

- **操作系统**：Windows 10 / Windows 11
- **Office 版本**：Microsoft 365 或 Office 2021 及以上版本
- **Excel 应用**：已安装且可正常打开
- **用户权限**：当前账户具有本机文件读写权限

> **提示**：本安装方法仅适用于 Windows 版 Office，macOS 版本不支持此方式。

#### 2. 下载清单文件

在浏览器中打开以下地址下载清单文件：

```
https://office-addin.bigmodel.cn/manifest.prod.xml
```

**下载后请确认：**
- [ ] 文件名称为 `manifest.prod.xml`
- [ ] 文件后缀为 `.xml`（非 `.txt`）

---

### 安装步骤

#### 第一步：放置清单文件到 Wef 文件夹

1. 按下键盘快捷键 `Win + R`，打开“运行”对话框。

2. 在输入框中粘贴以下路径，按回车确认：

   ```
   %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
   ```

   > **说明**：如果该路径下没有 `Wef` 文件夹，请手动创建。

3. 将下载好的 `manifest.prod.xml` 文件**直接复制**到 `Wef` 文件夹的根目录下。

   **注意事项：**
   - 不要放入子文件夹
   - 不要修改文件名
   - 确保文件后缀为 `.xml`

#### 第二步：共享 Wef 文件夹

1. 右键点击 `Wef` 文件夹，选择 **“属性”**。

2. 切换到 **“共享”** 选项卡。

3. 点击 **“共享(S)...”** 按钮。

4. 在用户列表中添加用户（建议添加 `Everyone`），并将权限级别设置为 **“读取/写入”**。

5. 点击 **“共享”** 完成设置。

6. 共享成功后，系统会显示网络路径，例如：

   ```
   \\YOUR_COMPUTER_NAME\Wef
   ```

   > **请记录此路径**，后续步骤将需要使用。

#### 第三步：在 Excel 中信任共享路径

1. 打开 Excel，依次点击：

   ```
   文件 → 选项 → 信任中心 → 信任中心设置(T)...
   ```

2. 在左侧菜单中，选择 **“受信任的加载项目录”**。

3. 在 **“目录 URL(U)”** 输入框中，粘贴上一步记录的网络路径（例如 `\\DESKTOP-XXXX\Wef`）。

4. 点击 **“添加目录(D)”**。

5. **勾选**新添加目录对应的 **“在菜单中显示”** 复选框。

6. 点击 **“确定”** 保存所有设置。

#### 第四步：加载插件

1. **重启 Excel**（必须步骤，使信任设置生效）。

2. 在 Excel 顶部菜单栏中点击 **“插入”**。

3. 点击 **“我的加载项”**。

4. 在弹出窗口的顶部，选择 **“共享文件夹”** 选项卡。

5. 在列表中找到您的插件，点击选中。

6. 点击 **“添加”** 按钮。

7. 插件加载完成后，即可在 Excel 右侧任务窗格中开始使用。

---

### 常见问题排查

| 问题现象 | 可能原因 | 解决方案 |
| :--- | :--- | :--- |
| “我的加载项”中看不到插件 | 信任路径未正确添加 | 重新检查第三步，确保路径与共享路径完全一致 |
| 加载时提示证书错误 | 自签名证书未受信任 | 确保已运行 `mkcert -install` |
| 插件加载但无法执行操作 | Wef 文件夹权限不足 | 确认共享权限为“读取/写入” |
| 任务窗格显示空白 | CSP 策略限制或资源加载失败 | 检查网络连接，按 F12 查看开发者工具 Console 报错 |
| 网络路径找不到 | 电脑名称变化或共享未开启 | 重新执行第二步，确认电脑名称及共享状态 |

#### 安装与信任路径问答

**Q：为什么“我的加载项”里看不到插件？**

A：最常见原因是信任路径未正确添加，或“目录 URL”与共享路径不一致。请重新执行第三步，确保路径与共享路径完全一致，并勾选“在菜单中显示”。

**Q：添加目录时提示找不到网络路径？**

A：`\\YOUR_COMPUTER_NAME\Wef` 中的 `YOUR_COMPUTER_NAME` 是占位符，需替换为实际电脑名称。可在“设置 → 系统 → 关于”中查看设备名称，或直接在文件资源管理器地址栏输入共享文件夹路径进行核对。

**Q：重启 Excel 后插件仍不显示？**

A：确认信任设置已保存并已完全重启 Excel；若仍不生效，请先完全退出 Excel（包括任务栏后台进程）再重新打开，并再次核对共享权限是否为“读取/写入”。

**Q：加载时提示证书错误？**

A：请确保已运行 `mkcert -install`，使自签名证书受信任。

**Q：插件加载了但无法执行操作？**

A：通常是 Wef 文件夹权限不足。请确认共享权限为“读取/写入”，而非“只读”。

**Q：多用户共用同一台电脑，需要分别配置吗？**

A：需要。每个用户需在自己的账户下分别完成第三步的信任设置。

**Q：找不到我的加载项怎么办？**

A：在设置中找到自定义功能区，选择所有命令，可找加载项。

---

### 注意事项

1. **测试用途声明**：此“共享文件夹”部署方式**仅适用于开发与测试**。微软官方**不支持**将其用于生产环境下的插件分发。

2. **平台限制**：此方法**仅适用于 Windows 版 Office**。macOS 用户需通过其他方式（如集中部署或商店发布）进行安装。

3. **更新机制**：如果插件更新涉及界面变化（如新增按钮或功能入口），用户可能需要**重新安装**插件才能看到变化。

4. **网络路径稳定**：确保电脑的网络名称（Computer Name）保持稳定，避免共享路径失效。

5. **多用户环境**：如果同一台电脑的多个用户需要使用，每个用户需分别执行信任步骤。

---

## 连接模型

### 方式一(推荐):API Key

对大多数用户来说,API Key 是最顺畅的方式,通常**无需**代理。

1. 在 Pi 中输入 `/login`(或使用欢迎页)
2. 展开某个服务商(OpenAI、Google Gemini、Anthropic 等)
3. 粘贴你的 API Key
4. 点击 **Save**

### 方式二:自定义 OpenAI 兼容网关(DeepSeek、智谱 GLM、本地模型等)

任何提供 OpenAI 兼容接口的服务都可以接入:

1. 在 Pi 中打开 `/settings`
2. 在 **Custom OpenAI-compatible gateways** 下填写:
   - **Endpoint**(接口基础地址 / base URL)
   - **Model**(模型 ID)
   - **API key**(部分本地服务可留空)
3. 保存网关后,在 `/model` 中选择该模型

常见示例(模型 ID 与地址请以各服务商官方文档为准):

| 服务商 | Endpoint | 模型 ID 示例 |
|---|---|---|
| DeepSeek | `https://api.deepseek.com` | `deepseek-chat`、`deepseek-reasoner` |
| 智谱 GLM(BigModel) | `https://open.bigmodel.cn/api/paas/v4` | `glm-4.6` 等 |
| Ollama(本地) | `http://localhost:11434/v1` | 本地已下载的模型 |

注意:

- 网关若是公网 HTTPS 地址,通常可直接连接,无需代理。
- localhost / 内网地址需经本地代理转发,启动 `pi4office-proxy` 时可能需要配置目标主机策略环境变量(如 `ALLOWED_TARGET_HOSTS`、`ALLOW_LOOPBACK_TARGETS`、`ALLOW_PRIVATE_TARGETS`),详见[英文安装指南](./docs/install.md#connect-a-provider)。

### 方式三:OAuth 账号登录

支持 Anthropic、OpenAI ChatGPT、Google Code Assist / Antigravity、GitHub Copilot。

1. 在 `/login` 中点击 **Login with …**
2. 在弹出的浏览器窗口中完成登录
3. 返回 Excel,按提示完成剩余步骤
   - OpenAI 与 Google 的 OAuth 流程中,浏览器最后会跳到一个显示**"无法访问此网站"**的页面——这是正常现象!复制浏览器地址栏中的完整 URL,粘贴回 Pi for Office 的提示框即可
   - 部分 Google Workspace 套餐还会要求填写 Google Cloud 项目 ID

#### OAuth 登录报 CORS / 网络错误?

Office 内嵌浏览器会拦截部分 OAuth 接口(典型报错:`Login was blocked by browser CORS`、`Load failed`、`Failed to fetch`)。解决方法是在本机运行一个本地 HTTPS 代理:

```bash
npx pi4office-proxy
```

(若未安装 Node.js:`curl -fsSL https://piforexcel.com/proxy | sh`)

然后在 Pi 中打开 `/settings` → **Proxy**,启用代理并填入代理启动时打印的 HTTPS 地址(通常是 `https://localhost:3003`; 如端口被占用,会显示另一个本地端口),重试登录。详细说明与排错见[英文安装指南](./docs/install.md#oauth-logins-and-cors-proxy)。API Key 方式一般不需要代理。

---

## 常见问题(简)

- **"我的加载项"里看不到插件** —— 重新检查信任路径是否添加正确,确保路径与共享路径完全一致;见上文"常见问题排查"
- **侧边栏打开但是空白** —— 检查网络连接,按 F12 查看开发者工具 Console 报错;见上文"常见问题排查"
- **如何更新** —— 大多数更新自动生效,关闭并重新打开侧边栏即可;若更新涉及界面变化,需重新安装插件

更多排错项见[英文安装指南 · Troubleshooting](./docs/install.md#troubleshooting)。

---

## 更多文档(英文)

| 文档 | 说明 |
|---|---|
| [docs/guide-zh-CN.md](./docs/guide-zh-CN.md) | **中文技术文档**——功能详解、使用方式、技术架构与管线 |
| [README.md](./README.md) | 完整功能介绍、开发者快速上手、架构说明 |
| [docs/install.md](./docs/install.md) | 完整安装指南 |
| [docs/integrations-external-tools.md](./docs/integrations-external-tools.md) | 联网搜索 + MCP 集成配置 |
| [docs/extensions.md](./docs/extensions.md) | 扩展开发指南 |
| [docs/security-threat-model.md](./docs/security-threat-model.md) | 安全威胁模型 |
| [docs/wps-support.md](./docs/wps-support.md) | WPS 表格支持现状与安装路径 |

## 许可证

[MIT](LICENSE) © Thomas Mustier
