#!/usr/bin/env node

/**
 * One-step local dev: ensure HTTPS certs, start the dev server, and
 * sideload the add-in into Excel and Word. Ctrl+C stops the server.
 *
 * Usage: npm run use
 */

import { spawn, spawnSync } from "node:child_process";
import fs from "node:fs";
import https from "node:https";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const CERT_PATH = path.join(ROOT_DIR, "cert.pem");
const KEY_PATH = path.join(ROOT_DIR, "key.pem");
const DEV_PORT = 3141;
const DEV_URL = `https://localhost:${DEV_PORT}`;
const SHELL = process.platform === "win32";

const log = (msg) => console.log(`[pi4office] ${msg}`);
const fail = (msg) => {
  console.error(`[pi4office] ${msg}`);
  process.exit(1);
};

function runSync(command, args) {
  const result = spawnSync(command, args, { cwd: ROOT_DIR, stdio: "inherit", shell: SHELL });
  if (result.error) {
    fail(`${command} is not available: ${result.error.message}`);
  }
  if (result.status !== 0) {
    fail(`Command failed: ${command} ${args.join(" ")} (exit ${result.status})`);
  }
}

function isServerUp() {
  return new Promise((resolve) => {
    const req = https.get(DEV_URL, { rejectUnauthorized: false }, (res) => {
      res.resume();
      resolve(true);
    });
    req.on("error", () => resolve(false));
    req.setTimeout(1000, () => {
      req.destroy();
      resolve(false);
    });
  });
}

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitForServer(timeoutMs = 60_000, child = null) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (child && child.exitCode !== null) {
      fail(
        `Dev server exited before becoming ready (code ${child.exitCode}). ` +
          `Check for a port conflict on ${DEV_PORT} or a build error above.`,
      );
    }
    if (await isServerUp()) return;
    await delay(500);
  }
  fail(`Dev server did not respond on ${DEV_URL} within ${timeoutMs / 1000}s. Check for a port conflict on ${DEV_PORT}.`);
}

function runSideload(app) {
  return new Promise((resolve) => {
    log(`Sideloading the add-in into ${app}...`);
    const child = spawn(
      "npx",
      ["office-addin-debugging", "start", "manifest.xml", "desktop", "--app", app, "--no-debug"],
      { cwd: ROOT_DIR, stdio: "inherit", shell: SHELL },
    );
    child.on("exit", (code) => {
      if (code !== 0) {
        fail(`Sideload into ${app} failed (exit ${code}).`);
      }
      resolve();
    });
  });
}

function killTree(child) {
  if (process.platform === "win32") {
    spawnSync("taskkill", ["/pid", String(child.pid), "/T", "/F"], { stdio: "ignore" });
  } else {
    child.kill("SIGTERM");
  }
}

// 1. HTTPS certificates — vite serves https://localhost:3141 via cert.pem/key.pem.
//    They are gitignored, so generate them with mkcert when missing.
if (!fs.existsSync(CERT_PATH) || !fs.existsSync(KEY_PATH)) {
  log("Missing HTTPS certificates (cert.pem / key.pem). Generating with mkcert...");
  runSync("mkcert", ["-install"]);
  runSync("mkcert", ["-cert-file", "cert.pem", "-key-file", "key.pem", "localhost"]);
}

// 2. Start the dev server, unless something already serves this URL.
let vite = null;
if (await isServerUp()) {
  log(`Dev server already responding on ${DEV_URL}; reusing it.`);
} else {
  log(`Starting dev server on ${DEV_URL}...`);
  vite = spawn("npm", ["run", "dev"], { cwd: ROOT_DIR, stdio: "inherit", shell: SHELL });
  await waitForServer(60_000, vite);
  log("Dev server is up.");
}

// 3. Sideload into Excel, then Word.
await runSideload("excel");
await runSideload("word");

log("Done. The add-in is sideloaded and ready in Excel and Word.");
log(`Keep this terminal open; Ctrl+C stops the dev server.`);

// 4. Keep the dev server alive until Ctrl+C (only when this script started it).
let stopping = false;
const stop = (code = 0) => {
  if (stopping) return;
  stopping = true;
  if (vite) {
    log("Stopping dev server...");
    killTree(vite);
  }
  process.exit(code);
};
process.on("SIGINT", () => stop(0));
process.on("SIGTERM", () => stop(0));
if (vite) {
  vite.on("exit", (code) => stop(code ?? 0));
}
