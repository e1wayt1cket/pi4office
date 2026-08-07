# Deploy hosted build on Vercel (maintainers)

Pi for Office’s taskpane is a static site built by Vite (`dist/`).

Vercel is a good default host because it’s free for OSS/hobby usage and handles HTTPS + caching well.

## One-time setup

1. Create a new Vercel project
2. Import `tmustier/pi4office`
3. Framework preset: **Vite** (or “Other”)
4. Build settings:
   - **Build Command:** `npm run build`
   - **Output Directory:** `dist`

This repo includes `vercel.json` with:
- `outputDirectory: dist`
- an `ignoreCommand` deploy policy (`node scripts/vercel-ignore-command.mjs`) for `main`, PR previews, and manual deploys
- `/proxy` rewrite to `/proxy.sh` (bootstrap script for `npx pi4office-proxy`)
- a header rule to disable caching for `/src/taskpane.html` to make updates propagate reliably
- an enforced `Content-Security-Policy` on `/src/taskpane.html` (Office.js + provider/auth endpoints + localhost proxy + Pyodide CDN host).

### `ignoreCommand` policy

Automatic deploy behavior is:
- **build** for `main`
- **build** for pull requests (`VERCEL_GIT_PULL_REQUEST_ID` is set)
- **build** for manual deploys (`VERCEL_GIT_COMMIT_REF` is unset)
- **skip** non-PR feature branch pushes

Regression coverage lives in `tests/vercel-ignore-command.test.mjs` (run via `npm run test:security`).

If a host-specific regression appears, temporary rollback is a single-header change:
`Content-Security-Policy` → `Content-Security-Policy-Report-Only`.

## Production URL

After deploy, you’ll have a production URL like:

- `https://<project>.vercel.app`

Keep this URL stable; it becomes the base URL used by `manifest.prod.xml`.

## Generate / update the production manifest

The dev manifest (`manifest.xml`) points at `https://localhost:3141`.

Generate the production manifest with the hosted base URL:

```bash
ADDIN_BASE_URL="https://<project>.vercel.app" npm run manifest:prod
```

This writes:
- `manifest.prod.xml` (repo root)
- `public/manifest.prod.xml` (so the hosted site can offer a one-click download at `/manifest.prod.xml`)

## Updates (automatic)

For most UI/behavior changes:
- deploy a new build to the same Vercel project
- users get the update automatically next time they open the taskpane

If a release requires a manifest change (rare):
- update and redistribute `manifest.prod.xml`
- users re-upload the manifest in Excel

---

## GitHub Pages mirror (for networks where *.vercel.app is unreachable)

The static build can also be served from GitHub Pages as a mirror for networks
that cannot reach `*.vercel.app` (for example mainland China).

- **URL:** `https://<owner>.github.io/<repo>/` — for this repo,
  `https://e1wayt1cket.github.io/pi4office/`.
- **Repo-side files already in place:**
  - `.github/workflows/deploy-pages.yml` — builds with `--base=/pi4office/`
    (GitHub Pages project sites live under a subpath) and deploys on every push to `main`.
  - `manifest.pages.xml` — sideload this instead of `manifest.prod.xml`;
    it points to the GitHub Pages URL.
- **One-time enable:** repo **Settings → Pages → Source: GitHub Actions**.
  After that, pushes to `main` deploy automatically.

**Limitations (GitHub Pages cannot set custom response headers):**
- No `Content-Security-Policy`, `Cache-Control`, or the OAuth callback rewrites
  that `vercel.json` provides on Vercel. Core add-in function is unaffected;
  security hardening is reduced.
- `*.github.io` may also be blocked on some mainland China networks — verify
  before relying on this mirror. In that case, a local server
  (`npm run serve:dist`) remains the fallback.

