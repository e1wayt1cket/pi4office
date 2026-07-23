#!/usr/bin/env node

/**
 * Build the Pi for Office Windows installer.
 *
 * This script:
 * 1. Builds the frontend (npm run build)
 * 2. Downloads Node.js portable for Windows
 * 3. Downloads mkcert for Windows
 * 4. Creates a self-contained installer directory
 * 5. Optionally compiles to exe via NSIS
 *
 * Usage:
 *   node scripts/build-exe-installer.mjs              # Create installer directory
 *   node scripts/build-exe-installer.mjs --exe         # Also attempt to create exe
 *   node scripts/build-exe-installer.mjs --node-url=... # Custom Node.js download URL
 */

import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { createWriteStream } from "node:fs";
import { pipeline } from "node:stream/promises";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const installerDir = path.join(rootDir, "pkg", "installer");
const stagingDir = path.join(installerDir, "staging");
const distDir = path.join(rootDir, "dist");

const NODE_VERSION = "22.19.0";
const NODE_DOWNLOAD_URL = process.env.NODE_DOWNLOAD_URL
  || `https://nodejs.org/dist/v${NODE_VERSION}/node-v${NODE_VERSION}-win-x64.zip`;

const MKCERT_DOWNLOAD_URL = process.env.MKCERT_DOWNLOAD_URL
  || "https://dl.filippo.io/mkcert/latest?for=windows/amd64";

function log(msg) {
  console.log(`[build-installer] ${msg}`);
}

function run(cmd, args, opts = {}) {
  log(`Running: ${cmd} ${args.join(" ")}`);
  const result = spawnSync(cmd, args, {
    stdio: "inherit",
    cwd: rootDir,
    shell: true,
    ...opts,
  });
  if (result.error) {
    console.error(`[build-installer] Command failed: ${cmd}`);
    console.error(result.error.message);
    process.exit(1);
  }
  if (result.status !== 0 && !opts.ignoreError) {
    console.error(`[build-installer] Command exited with ${result.status}`);
    process.exit(result.status);
  }
  return result;
}

async function downloadFile(url, destPath) {
  log(`Downloading: ${url}`);
  const res = await fetch(url, { redirect: "follow" });
  if (!res.ok) {
    throw new Error(`Failed to download ${url}: ${res.status} ${res.statusText}`);
  }
  const fileStream = createWriteStream(destPath);
  await pipeline(res.body, fileStream);
  log(`Saved to: ${destPath}`);
}

function findMakensis() {
  // Check PATH first
  const pathCheck = spawnSync("makensis", ["-VERSION"], { stdio: "ignore", shell: true });
  if (pathCheck.status === 0) return "makensis";

  // Windows default install path
  const nsisPath = "C:\\Program Files (x86)\\NSIS\\makensis.exe";
  if (fs.existsSync(nsisPath)) return nsisPath;

  return null;
}

// --- Step 1: Build the frontend ---
function buildFrontend() {
  log("Building frontend...");
  if (!fs.existsSync(path.join(rootDir, "node_modules"))) {
    log("Installing dependencies...");
    run("npm", ["ci"]);
  }
  run("npm", ["run", "build"]);
  log("Frontend build complete.");
}

// --- Step 2: Download Node.js portable ---
async function downloadNode() {
  const nodeDir = path.join(stagingDir, "nodejs");
  if (fs.existsSync(nodeDir)) {
    log("Node.js portable already downloaded, skipping.");
    return;
  }

  const zipPath = path.join(stagingDir, "node.zip");
  fs.mkdirSync(nodeDir, { recursive: true });

  await downloadFile(NODE_DOWNLOAD_URL, zipPath);

  log("Extracting Node.js...");

  // Use PowerShell to extract zip (available on all Windows)
  run("powershell", [
    "-Command",
    `Expand-Archive -Path '${zipPath}' -DestinationPath '${stagingDir}' -Force`,
  ]);

  // Find the extracted directory (node-vX.Y.Z-win-x64)
  const extracted = fs.readdirSync(stagingDir)
    .find((f) => f.startsWith("node-v") && f.includes("win-x64"));

  if (extracted) {
    const extractedDir = path.join(stagingDir, extracted);
    // Move contents to nodejs/
    for (const file of fs.readdirSync(extractedDir)) {
      fs.renameSync(
        path.join(extractedDir, file),
        path.join(nodeDir, file),
      );
    }
    fs.rmdirSync(extractedDir);
  }

  fs.unlinkSync(zipPath);
  log("Node.js portable ready.");
}

// --- Step 3: Download mkcert ---
async function downloadMkcert() {
  const mkcertPath = path.join(stagingDir, "mkcert.exe");
  if (fs.existsSync(mkcertPath)) {
    log("mkcert already downloaded, skipping.");
    return;
  }

  fs.mkdirSync(stagingDir, { recursive: true });
  await downloadFile(MKCERT_DOWNLOAD_URL, mkcertPath);
  log("mkcert downloaded.");
}

// --- Step 4: Copy application files ---
function copyAppFiles() {
  log("Copying application files...");

  // Copy dist/
  const stagingDist = path.join(stagingDir, "dist");
  if (fs.existsSync(stagingDist)) {
    fs.rmSync(stagingDist, { recursive: true });
  }
  fs.cpSync(distDir, stagingDist, { recursive: true });
  log("Copied dist/");

  // Copy server script
  fs.copyFileSync(
    path.join(installerDir, "server.mjs"),
    path.join(stagingDir, "server.mjs"),
  );
  log("Copied server.mjs");

  // Copy install script
  fs.copyFileSync(
    path.join(installerDir, "install.ps1"),
    path.join(stagingDir, "install.ps1"),
  );
  log("Copied install.ps1");

  // Write a simple launcher batch file
  const batContent = `@echo off
title Pi for Office Server
echo Starting Pi for Office server...
echo.
"${path.join(stagingDir, "nodejs", "node.exe")}" "${path.join(stagingDir, "server.mjs")}"
if errorlevel 1 (
  echo.
  echo Server stopped with an error. Press any key to close.
  pause > nul
)
`;
  fs.writeFileSync(path.join(stagingDir, "start-server.bat"), batContent);
  log("Created start-server.bat");
}

// --- Step 5: Create NSIS exe (optional) ---
function buildNsisExe() {
  const makensis = findMakensis();
  if (!makensis) {
    log("NSIS (makensis) not found. Install with: winget install NSIS.NSIS");
    log("Alternative: use the staging/ directory directly with install.ps1");
    return false;
  }

  log("Building NSIS installer exe...");
  const nsiScript = path.join(installerDir, "setup.nsi");

  if (!fs.existsSync(nsiScript)) {
    log("NSIS script not found, skipping exe build.");
    log("The installer files are ready in: " + stagingDir);
    return false;
  }

  // Copy NSIS script + LICENSE to staging for relative-path resolution
  fs.copyFileSync(nsiScript, path.join(stagingDir, "setup.nsi"));
  const licensePath = path.join(rootDir, "LICENSE");
  if (fs.existsSync(licensePath)) {
    fs.copyFileSync(licensePath, path.join(stagingDir, "LICENSE"));
  }

  run(makensis, ["setup.nsi"], { cwd: stagingDir });
  log("Installer exe built successfully.");
  return true;
}

// --- Step 6: Create ZIP archive ---
function createZipArchive() {
  log("Creating portable ZIP archive...");
  const zipDest = path.join(installerDir, "pi4office-portable.zip");

  run("powershell", [
    "-Command",
    `Compress-Archive -Path '${stagingDir}\\*' -DestinationPath '${zipDest}' -Force`,
  ]);

  log(`ZIP archive: ${zipDest}`);
}

// --- Main ---
async function main() {
  const args = process.argv.slice(2);
  const buildExe = args.includes("--exe");
  const skipBuild = args.includes("--skip-build");
  const skipNode = args.includes("--skip-node");
  const skipMkcert = args.includes("--skip-mkcert");

  console.log("");
  log("=== Pi for Office Installer Builder ===");
  console.log("");

  // Clean staging
  if (fs.existsSync(stagingDir)) {
    fs.rmSync(stagingDir, { recursive: true });
  }
  fs.mkdirSync(stagingDir, { recursive: true });

  if (!skipBuild) {
    buildFrontend();
  } else {
    log("Skipping frontend build (--skip-build)");
  }

  if (!skipNode) {
    await downloadNode();
  } else {
    log("Skipping Node.js download (--skip-node)");
  }

  if (!skipMkcert) {
    await downloadMkcert();
  } else {
    log("Skipping mkcert download (--skip-mkcert)");
  }

  copyAppFiles();
  createZipArchive();

  if (buildExe) {
    const exeBuilt = buildNsisExe();
    if (!exeBuilt) {
      log("Exe build skipped, but ZIP archive is ready.");
      log(`Portable installer: ${installerDir}\\pi4office-portable.zip`);
    }
  } else {
    log("Skipping exe build (use --exe to attempt NSIS compilation).");
    log(`Portable installer: ${installerDir}\\pi4office-portable.zip`);
  }

  console.log("");
  log("=== Build complete ===");
  log(`Staging directory: ${stagingDir}`);
  log("To test the installer, run:");
  log(`  powershell -ExecutionPolicy Bypass -File "${stagingDir}\\install.ps1"`);
  console.log("");
}

main().catch((err) => {
  console.error("[build-installer] Fatal error:", err.message);
  process.exit(1);
});
