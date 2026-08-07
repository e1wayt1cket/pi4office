#!/usr/bin/env node

/**
 * Combined HTTPS server for Pi for Office local deployment.
 *
 * Serves the add-in frontend (dist/) and handles CORS proxy requests
 * on a single port. Designed for the self-contained Windows installer.
 *
 * Usage:
 *   node server.mjs [--port PORT] [--host HOST]
 */

import https from "node:https";
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Readable } from "node:stream";
import { lookup as dnsLookup } from "node:dns/promises";

const serverDir = path.dirname(fileURLToPath(import.meta.url));

function parseArgs() {
  const args = process.argv.slice(2);
  let host = "localhost";
  let port = 3141;
  let certDir = path.join(serverDir, "certs");

  for (let i = 0; i < args.length; i++) {
    if (args[i] === "--port" && args[i + 1]) {
      port = Number.parseInt(args[i + 1], 10) || port;
      i++;
    } else if (args[i] === "--host" && args[i + 1]) {
      host = args[i + 1];
      i++;
    } else if (args[i] === "--cert-dir" && args[i + 1]) {
      certDir = args[i + 1];
      i++;
    }
  }

  return { host, port, certDir };
}

const { host: HOST, port: PORT, certDir: CERT_DIR } = parseArgs();
const DIST_DIR = path.join(serverDir, "dist");
const KEY_PATH = process.env.TLS_KEY_PATH || path.join(CERT_DIR, "key.pem");
const CERT_PATH = process.env.TLS_CERT_PATH || path.join(CERT_DIR, "cert.pem");

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".woff2": "font/woff2",
  ".woff": "font/woff",
  ".ttf": "font/ttf",
  ".txt": "text/plain; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
};

const HOP_BY_HOP_HEADERS = new Set([
  "connection", "keep-alive", "proxy-authenticate", "proxy-authorization",
  "te", "trailer", "transfer-encoding", "upgrade",
]);

const DEFAULT_ALLOWED_ORIGINS = new Set([
  "https://localhost:3141",
  "https://pi4office.vercel.app",
]);

const allowedOrigins = (() => {
  const raw = process.env.ALLOWED_ORIGINS;
  if (!raw) return DEFAULT_ALLOWED_ORIGINS;
  const set = new Set(raw.split(",").map((s) => s.trim()).filter(Boolean));
  return set.size > 0 ? set : DEFAULT_ALLOWED_ORIGINS;
})();

const DEFAULT_ALLOWED_TARGET_HOSTS = new Set([
  "api.anthropic.com", "console.anthropic.com", "platform.claude.com",
  "github.com", "api.github.com",
  "auth.openai.com", "api.openai.com", "chatgpt.com",
  "oauth2.googleapis.com", "generativelanguage.googleapis.com",
  "cloudcode-pa.googleapis.com", "daily-cloudcode-pa.sandbox.googleapis.com",
  "api.z.ai",
  "s.jina.ai", "api.firecrawl.dev", "google.serper.dev",
  "api.tavily.com", "api.search.brave.com",
]);

const allowAllTargetHosts = ["1", "true"].includes(process.env.ALLOW_ALL_TARGET_HOSTS || "");
const allowLoopbackTargets = ["1", "true"].includes(process.env.ALLOW_LOOPBACK_TARGETS || "");
const allowPrivateTargets = ["1", "true"].includes(process.env.ALLOW_PRIVATE_TARGETS || "");

const configuredHosts = (process.env.ALLOWED_TARGET_HOSTS || "").trim();
const configuredAllowedTargetHosts = configuredHosts
  ? new Set(configuredHosts.split(",").map((s) => s.trim()).filter(Boolean))
  : null;

const allowedTargetHosts = (() => {
  if (allowAllTargetHosts) return new Set();
  if (configuredAllowedTargetHosts && configuredAllowedTargetHosts.size > 0) return configuredAllowedTargetHosts;
  return new Set(DEFAULT_ALLOWED_TARGET_HOSTS);
})();

function isLoopbackAddress(addr) {
  if (!addr) return false;
  if (addr === "::1" || addr === "0:0:0:0:0:0:0:1") return true;
  if (addr.startsWith("127.")) return true;
  if (addr.startsWith("::ffff:127.")) return true;
  return false;
}

function isIpLiteral(host) {
  if (!host) return false;
  if (host.includes(":") && !host.startsWith("[")) {
    try { new URL(`http://[${host}]`); return true; } catch { return false; }
  }
  try { new URL(`http://${host}`); } catch { return false; }
  return host.split(".").length === 4 && host.split(".").every((p) => /^\d+$/.test(p) && Number(p) <= 255);
}

function isPrivateIp(ip) {
  if (ip.startsWith("10.") || ip.startsWith("192.168.")) return true;
  if (ip.startsWith("172.")) {
    const second = Number.parseInt(ip.split(".")[1], 10);
    if (second >= 16 && second <= 31) return true;
  }
  return false;
}

function safeJoin(base, reqPath) {
  const decoded = decodeURIComponent(reqPath);
  const cleaned = decoded.replace(/^\/+/, "");
  const full = path.resolve(base, cleaned);
  const rel = path.relative(base, full);
  if (rel.startsWith("..") || path.isAbsolute(rel)) {
    throw new Error("Path traversal");
  }
  return full;
}

function setCorsHeaders(req, res) {
  const origin = req.headers.origin;
  if (origin && allowedOrigins.has(origin)) {
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Vary", "Origin");
  }
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,PATCH,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers",
    req.headers["access-control-request-headers"] || "*");
  res.setHeader("Access-Control-Expose-Headers",
    "*, x-pi4office-proxy, x-pi4office-codex-websocket-bridge");
  res.setHeader("Access-Control-Max-Age", "86400");
}

function buildOutboundHeaders(inHeaders) {
  const out = new Headers();
  for (const [key, value] of Object.entries(inHeaders)) {
    if (!value) continue;
    const lower = key.toLowerCase();
    if (lower === "host" || lower === "content-length" || lower === "accept-encoding") continue;
    if (lower === "user-agent" || lower === "accept-language") continue;
    if (lower === "origin" || lower === "referer") continue;
    if (lower.startsWith("sec-fetch-") || lower.startsWith("sec-ch-")) continue;
    if (lower === "anthropic-dangerous-direct-browser-access") continue;
    if (lower === "cookie") continue;
    if (HOP_BY_HOP_HEADERS.has(lower)) continue;
    if (Array.isArray(value)) {
      for (const v of value) out.append(key, v);
    } else {
      out.set(key, value);
    }
  }
  return out;
}

function extractTargetUrl(rawUrl) {
  const idx = rawUrl.indexOf("url=");
  if (idx === -1) return null;
  const encoded = rawUrl.slice(idx + 4);
  const normalized = encoded.replace(/\+/g, "%20");
  try {
    return decodeURIComponent(normalized);
  } catch {
    return null;
  }
}

async function handleProxyRequest(req, res) {
  const origin = req.headers.origin;
  if (!origin || !allowedOrigins.has(origin)) {
    res.statusCode = 403;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("forbidden");
    return;
  }

  setCorsHeaders(req, res);

  if (req.method === "OPTIONS") {
    res.statusCode = 204;
    res.end();
    return;
  }

  const rawUrl = req.url || "/";

  // Health check
  if (rawUrl.split("?")[0] === "/healthz") {
    res.statusCode = 200;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.setHeader("x-pi4office-proxy", "1");
    res.setHeader("x-pi4office-codex-websocket-bridge", "1");
    res.end("ok");
    return;
  }

  const target = extractTargetUrl(rawUrl);
  if (!target) {
    res.statusCode = 400;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("Missing or invalid ?url=<target-url> query parameter");
    return;
  }

  let targetUrl;
  try { targetUrl = new URL(target); } catch {
    res.statusCode = 400;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("Invalid target URL");
    return;
  }

  if (targetUrl.protocol !== "http:" && targetUrl.protocol !== "https:") {
    res.statusCode = 400;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("Only http(s) target URLs are supported");
    return;
  }

  const targetHost = targetUrl.hostname;
  const safeTarget = `${targetUrl.origin}${targetUrl.pathname}`;

  // Security: validate target host
  if (!allowAllTargetHosts && allowedTargetHosts.size > 0) {
    if (!isIpLiteral(targetHost)) {
      if (!allowedTargetHosts.has(targetHost)) {
        res.statusCode = 403;
        res.setHeader("Content-Type", "text/plain; charset=utf-8");
        res.end("Target host is not allowlisted");
        console.warn(`[proxy] blocked target: ${safeTarget}`);
        return;
      }
    }
  }

  try {
    const startedAt = Date.now();
    const headers = buildOutboundHeaders(req.headers);
    const hasBody = req.method && !["GET", "HEAD"].includes(req.method);
    const body = hasBody ? Readable.toWeb(req) : undefined;

    const upstream = await fetch(targetUrl.toString(), {
      method: req.method,
      headers,
      body,
      ...(body ? { duplex: "half" } : {}),
      redirect: "manual",
    });

    console.log(`[proxy] ${req.method || "GET"} ${safeTarget} -> ${upstream.status} (${Date.now() - startedAt}ms)`);

    res.statusCode = upstream.status;

    upstream.headers.forEach((value, key) => {
      const lower = key.toLowerCase();
      if (lower === "set-cookie") return;
      if (HOP_BY_HOP_HEADERS.has(lower)) return;
      if (lower === "content-encoding" || lower === "content-length") return;
      if (lower.startsWith("access-control-") || lower === "vary") return;
      res.setHeader(key, value);
    });

    if (!upstream.body) {
      res.end();
      return;
    }

    const nodeStream = Readable.fromWeb(upstream.body);
    nodeStream.on("error", () => { try { res.end(); } catch {} });
    nodeStream.pipe(res);
  } catch (err) {
    console.warn(`[proxy] ${req.method || "GET"} ${safeTarget} -> ERROR (${err instanceof Error ? err.message : String(err)})`);
    res.statusCode = 502;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end(`Proxy error: ${err instanceof Error ? err.message : String(err)}`);
  }
}

function serveStaticFile(req, res) {
  try {
    const remote = req.socket?.remoteAddress;
    if (!isLoopbackAddress(remote)) {
      res.statusCode = 403;
      res.setHeader("Content-Type", "text/plain; charset=utf-8");
      res.end("forbidden");
      return;
    }

    const url = new URL(req.url || "/", `https://${HOST}:${PORT}`);
    let reqPath = url.pathname;
    if (reqPath === "/") reqPath = "/src/taskpane.html";

    const filePath = safeJoin(DIST_DIR, reqPath);
    if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
      res.statusCode = 404;
      res.setHeader("Content-Type", "text/plain; charset=utf-8");
      res.end("not found");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    res.statusCode = 200;
    res.setHeader("Content-Type", MIME[ext] || "application/octet-stream");

    if (reqPath.startsWith("/assets/") && /-[A-Za-z0-9]{8,}\./.test(reqPath)) {
      res.setHeader("Cache-Control", "public, max-age=31536000, immutable");
    } else {
      res.setHeader("Cache-Control", "no-cache");
    }

    fs.createReadStream(filePath).pipe(res);
  } catch (e) {
    res.statusCode = 500;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end(`server error: ${e instanceof Error ? e.message : String(e)}`);
  }
}

function createServer() {
  if (!fs.existsSync(KEY_PATH) || !fs.existsSync(CERT_PATH)) {
    console.error("[pi4office-server] Missing TLS certificates.");
    console.error(`  key:  ${KEY_PATH}${fs.existsSync(KEY_PATH) ? "" : " (missing)"}`);
    console.error(`  cert: ${CERT_PATH}${fs.existsSync(CERT_PATH) ? "" : " (missing)"}`);
    console.error("[pi4office-server] Re-run the installer to generate certificates.");
    process.exit(1);
  }

  return https.createServer(
    {
      key: fs.readFileSync(KEY_PATH),
      cert: fs.readFileSync(CERT_PATH),
    },
    (req, res) => {
      const reqPath = (req.url || "/").split("?")[0];

      // Proxy routes: /api-proxy/...
      if (reqPath.startsWith("/api-proxy/")) {
        // Rewrite URL to strip the prefix for the proxy handler
        req.url = req.url.replace(/^\/api-proxy/, "");
        handleProxyRequest(req, res);
        return;
      }

      // Static file serving
      serveStaticFile(req, res);
    },
  );
}

function start() {
  if (!fs.existsSync(DIST_DIR)) {
    console.error(`[pi4office-server] dist/ not found at ${DIST_DIR}`);
    console.error("[pi4office-server] Run the installer to set up the application files.");
    process.exit(1);
  }

  const server = createServer();
  server.listen(PORT, () => {
    console.log(`[pi4office-server] Pi for Office running at https://${HOST}:${PORT}`);
    console.log(`[pi4office-server] Serving: ${DIST_DIR}`);
    console.log(`[pi4office-server] Proxy: https://${HOST}:${PORT}/api-proxy/?url=<target>`);
    console.log("[pi4office-server] Press Ctrl+C to stop");
  });

  server.on("error", (err) => {
    if (err.code === "EADDRINUSE") {
      console.error(`[pi4office-server] Port ${PORT} is already in use.`);
      console.error(`[pi4office-server] Stop any running pi4office server and try again.`);
    } else {
      console.error(`[pi4office-server] Server error: ${err.message}`);
    }
    process.exit(1);
  });
}

start();
