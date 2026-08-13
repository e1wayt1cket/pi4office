# Install Pi for Office

> 中文用户:简要中文安装与模型配置指南见 [README.zh-CN.md](../README.zh-CN.md)。

## Prerequisites

### 1. Confirm system requirements

Before you start, make sure your environment meets the following requirements:

- **Operating system**: Windows 10 / Windows 11
- **Office version**: Microsoft 365 or Office 2021 or later
- **Excel**: installed and can be opened normally
- **User permissions**: your account has read/write access to local files

> **Note**: This installation method only applies to Windows Office. macOS does not support this method.

### 2. Download the manifest file

Open the following address in your browser to download the manifest file:

```
https://office-addin.bigmodel.cn/manifest.prod.xml
```

**After downloading, confirm:**
- [ ] The file is named `manifest.prod.xml`
- [ ] The file extension is `.xml` (not `.txt`)

---

## Installation

### Step 1: Place the manifest file in the Wef folder

1. Press `Win + R` to open the Run dialog.

2. Paste the following path into the input box and press Enter:

   ```
   %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
   ```

   > **Note**: If there is no `Wef` folder at this path, create it manually.

3. Copy the downloaded `manifest.prod.xml` file **directly** into the root of the `Wef` folder.

   **Notes:**
   - Do not put it in a subfolder
   - Do not rename the file
   - Make sure the file extension is `.xml`

### Step 2: Share the Wef folder

1. Right-click the `Wef` folder and select **Properties**.

2. Switch to the **Sharing** tab.

3. Click **Share…**.

4. Add users to the list (adding `Everyone` is recommended) and set the permission level to **Read/Write**.

5. Click **Share** to finish.

6. After sharing succeeds, a network path is displayed, for example:

   ```
   \\YOUR_COMPUTER_NAME\Wef
   ```

   > **Record this path** — you will need it in the next step.

### Step 3: Trust the shared path in Excel

1. Open Excel and click through:

   ```
   File → Options → Trust Center → Trust Center Settings…
   ```

2. In the left menu, select **Trusted Add-in Catalogs**.

3. In the **Catalog URL** box, paste the network path from the previous step (for example `\\DESKTOP-XXXX\Wef`).

4. Click **Add catalog**.

5. **Tick** the **Show in Menu** checkbox for the newly added catalog.

6. Click **OK** to save all settings.

### Step 4: Load the add-in

1. **Restart Excel** (required for the trust settings to take effect).

2. In the Excel top menu bar, click **Insert**.

3. Click **My Add-ins**.

4. At the top of the dialog, select the **Shared Folder** tab.

5. Find your add-in in the list and select it.

6. Click **Add**.

7. Once the add-in has loaded, you can start using it in the task pane on the right side of Excel.

---

## Troubleshooting

| Symptom | Possible cause | Solution |
| :--- | :--- | :--- |
| Add-in not visible in My Add-ins | Trusted path was not added correctly | Re-check step 3 and make sure the path exactly matches the shared path |
| Certificate error when loading | Self-signed certificate is not trusted | Make sure `mkcert -install` has been run |
| Add-in loads but cannot perform actions | Insufficient Wef folder permissions | Confirm the share permission is Read/Write |
| Task pane is blank | CSP policy restriction or resource load failure | Check the network connection and press F12 to view Console errors in the developer tools |
| Network path not found | Computer name changed or sharing is not enabled | Repeat step 2 and confirm the computer name and sharing state |

---

## Notes

1. **Testing note**: this Shared Folder deployment method is **only for development and testing**. Microsoft officially does **not** support using it to distribute add-ins in production.

2. **Platform limit**: this method **only applies to Windows Office**. macOS users must install another way (such as centralized deployment or store publishing).

3. **Updates**: if an add-in update changes the UI (such as new buttons or entry points), users may need to **reinstall** the add-in to see the change.

4. **Stable network path**: keep the computer's network name (Computer Name) stable so the shared path does not stop working.

5. **Multiple users**: if multiple users on the same computer need to use the add-in, each user must run the trust steps separately.

---

## First-run check

1. Open the taskpane (click the **Add-ins** button in the Home ribbon tab, then click **Pi for Office**)
2. Connect a provider (see below)
3. Send a test prompt, e.g.:
   - `What sheet am I currently on?`
   - `Summarize my current selection`

If you get a response, install is complete.

---

## Connect a provider

### Recommended (easiest): API key

For most users, API keys are the smoothest setup and usually do **not** need the proxy.

1. In Pi, run `/login` (or use the welcome screen)
2. Expand a provider row (OpenAI, Google Gemini, Anthropic, etc.)
3. Paste your API key
4. Click **Save**

### Custom OpenAI-compatible gateway (company or local)

Use this when your org exposes an OpenAI-compatible endpoint (or for local OpenAI-compatible servers).

1. In Pi, open `/settings`
2. Under **Custom OpenAI-compatible gateways**, set:
   - **Endpoint** (base URL)
   - **Model** (model ID)
   - **API key** (optional for some local servers)
3. Save the gateway, then choose its model from `/model`

Notes:
- If your gateway is publicly reachable over HTTPS, you can usually connect directly (no proxy).
- For localhost/private endpoints via the local proxy, you may need to configure proxy host policy env vars (for example `ALLOWED_TARGET_HOSTS`, `ALLOW_LOOPBACK_TARGETS`, or `ALLOW_PRIVATE_TARGETS`) when starting `pi4office-proxy`.

### OAuth / account login (Anthropic, OpenAI ChatGPT, Google Code Assist/Antigravity, GitHub Copilot)

1. In `/login`, click **Login with …**
2. Complete login in the browser window that opens
3. Return to Excel and complete any prompt shown
   - With the local proxy running, ChatGPT, Anthropic, and Google OAuth should continue automatically after the browser redirects to localhost.
   - If automatic capture is unavailable, your browser may land on a page that says **"can't be reached"** — that's normal! Copy the full URL from the browser address bar and paste it when prompted in Pi for Office.
   - Some Google workspace tiers may also ask for a Google Cloud project ID during setup

If login fails with a CORS/network error, follow the next section.

---

## OAuth logins and CORS proxy

Some OAuth/token endpoints are blocked by CORS inside Office webviews (especially on macOS WKWebView).

Typical symptoms:
- `Login was blocked by browser CORS`
- `Load failed`
- `Failed to fetch`

> **Managed / org rollout?** Instead of a proxy on every machine, IT can host one central proxy for everyone — see [Org-hosted central CORS proxy](./central-proxy.md).

### What to do

1. Run a local HTTPS proxy on the same machine as Excel (defaults to `https://localhost:3003`):

> ⚠️ **You may be asked for your Mac password** during this step. The proxy creates a local security certificate so Excel can talk to it securely. This is a one-time setup. If you are not an admin on this machine, ask your IT team to run this step for you.

If you already have Node.js:

```bash
npx pi4office-proxy
```

If you do not have Node.js (or are unsure):

```bash
curl -fsSL https://piforexcel.com/proxy | sh
```

2. In Pi, open `/settings` → **Proxy**:
   - enable **Proxy**
   - set URL to the URL printed by the proxy (normally `https://localhost:3003`; if 3003 is busy for another service, it will choose a random free port and print that URL)

3. Retry OAuth login

Quick proxy sanity check (advanced):
- In Terminal, run:

```bash
curl -k -i -s \
  'https://localhost:3141/api-proxy/google-cloudcode/v1internal:streamGenerateContent?alt=sse' \
  -X POST -H 'content-type: application/json' -d '{}' | head
```

- `401` means proxy routing is working (request reached Google, but without auth token).
- `404` usually means a proxy/path issue.
- Use single quotes around the URL in zsh so `?alt=sse` is not treated as a glob.

Notes:
- Keep the proxy URL on **HTTPS** (`https://...`), not HTTP.
- API-key providers generally work without proxy.
- The local proxy also starts loopback-only callback listeners for browser OAuth flows so ChatGPT (`http://localhost:1455/auth/callback`), Anthropic (`http://localhost:53692/callback`), Google Code Assist (`http://localhost:8085/oauth2callback`), and Google Antigravity (`http://localhost:51121/oauth-callback`) can capture browser callbacks automatically. If a port is busy, the affected login still works via the manual URL paste fallback.
- `3141` is the add-in/dev-server port; the local proxy normally uses `3003`.
- GPT-5.6 Luna on the ChatGPT provider requires the current proxy's Codex WebSocket bridge. `https://localhost:3003/healthz` must advertise both `x-pi4office-proxy: 1` and `x-pi4office-codex-websocket-bridge: 1`.
- If an older proxy is running, stop that process and rerun `npx -y pi4office-proxy@latest`; the current CLI intentionally refuses to reuse an outdated listener.
- If a healthy compatible proxy is already running on `3003`, the CLI reports that and exits instead of starting a duplicate.
- If port `3003` is busy for some other reason, the CLI automatically chooses a random free port. Copy the printed `https://localhost:<port>` URL into `/settings` → **Proxy**.
- To force a specific port, set `PORT` and use that same URL in settings:

```bash
PORT=3005 npx pi4office-proxy
```

---

## Updates

Pi for Office loads from a hosted URL, so most updates are automatic.

- Normal case: close and reopen the Excel taskpane to pick up the latest version.
- Rare case (manifest changes): download the new `manifest.prod.xml` and re-copy it into the Wef folder.

---

## Further troubleshooting

### Windows says the manifest certificate is invalid / mentions XML Expansion Packs
- `manifest.prod.xml` is an Office add-in manifest, not a legacy Excel XML Expansion Pack
- If you already tried the XML Expansion Packs path, close Excel and repeat the shared-folder flow above

### Do I need to install a separate Office.js bridge?
- No — Office.js support comes from Excel itself when you install Pi with `manifest.prod.xml`
- You do **not** need `generator-office`, Yeoman, or any extra Office.js package to use the hosted add-in
- The optional local helper services are only for OAuth proxying, native Python / LibreOffice, and tmux

### OAuth login still fails
- Confirm proxy is running and reachable at the exact URL in `/settings`
- Confirm proxy URL is `https://localhost:<port>` (not `http://`)
- Try API key auth as a fallback

---

## Developer setup (separate)

If you want to run from source (`localhost`, Vite, mkcert), use the root README: [Developer Quick Start](../README.md#developer-quick-start).
