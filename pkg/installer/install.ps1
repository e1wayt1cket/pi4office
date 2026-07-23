# Pi for Office Windows Installer
#
# Self-contained installer that sets up a local HTTPS server for the add-in,
# generates TLS certs, and registers the manifest with Excel.
#
# Run: powershell -ExecutionPolicy Bypass -File install.ps1

param(
  [switch]$Uninstall,
  [string]$InstallDir = "$env:LOCALAPPDATA\pi4office",
  [int]$Port = 3141
)

$ErrorActionPreference = "Stop"
$APP_NAME = "Pi for Office"
$APP_DIR = $InstallDir
$CERT_DIR = "$APP_DIR\certs"
$DIST_DIR = "$APP_DIR\dist"
$MANIFEST_FILE = "$APP_DIR\manifest.xml"
$SERVER_SCRIPT = "$APP_DIR\server.mjs"
$NODE_PORTABLE_DIR = "$APP_DIR\nodejs"
$NODE_EXE = "$NODE_PORTABLE_DIR\node.exe"
$MKCERT_EXE = "$APP_DIR\mkcert.exe"
$SHORTCUT_PATH = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Pi for Office.lnk"

$MANIFEST_CONTENT = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xsi:type="TaskPaneApp">
  <Id>a1b2c3d4-e5f6-7890-abcd-ef1234567890</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Pi for Office</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Pi for Office" />
  <Description DefaultValue="Open-source, multi-model AI assistant for Microsoft Office." />
  <IconUrl DefaultValue="https://localhost:3141/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://localhost:3141/assets/icon-80.png" />
  <SupportUrl DefaultValue="https://github.com/e1wayt1cket/pi4office" />
  <AppDomains>
    <AppDomain>https://localhost:3141</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Document" />
    <Host Name="Presentation" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3141/src/taskpane.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>600</RequestedHeight>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
"@

function Write-Step($msg) {
  Write-Host "  [pi4office] $msg" -ForegroundColor Cyan
}

function Write-Check($msg) {
  Write-Host "  [pi4office] + $msg" -ForegroundColor Green
}

function Write-Warn($msg) {
  Write-Host "  [pi4office] ! $msg" -ForegroundColor Yellow
}

function Write-Err($msg) {
  Write-Host "  [pi4office] x $msg" -ForegroundColor Red
}

function Get-OfficeWefDir {
  $wefPaths = @()

  # Microsoft Store Office path
  $storeWef = "$env:LOCALAPPDATA\Packages\Microsoft.Office.Desktop_8wekyb3d8bbwe\LocalCache\Content\Microsoft\WEF"
  if (Test-Path $storeWef) {
    $wefPaths += @{ App = "Office (Store)"; Path = $storeWef }
  }

  # Classic Win32 Office — top-level WEF
  $classicWef = "$env:LOCALAPPDATA\Microsoft\Office\WEF"
  if (-not (Test-Path $classicWef)) {
    New-Item -ItemType Directory -Path $classicWef -Force | Out-Null
  }
  $wefPaths += @{ App = "Office"; Path = $classicWef }

  # Version-specific path (e.g. 16.0\Wef) — required by some Office builds
  $officeRoot = "$env:LOCALAPPDATA\Microsoft\Office"
  if (Test-Path $officeRoot) {
    Get-ChildItem $officeRoot -Directory -Depth 0 -ErrorAction SilentlyContinue | ForEach-Object {
      $versionWef = Join-Path $_.FullName "Wef"
      if ((Test-Path $versionWef) -and $versionWef -ne $classicWef) {
        $wefPaths += @{ App = "Office $($_.Name)"; Path = $versionWef }
      }
    }
  }

  return $wefPaths
}

function Invoke-Uninstall {
  Write-Host "`n  === Pi for Office Uninstaller ===" -ForegroundColor Yellow
  Write-Host ""

  # Stop running server
  Get-Process -Name "node" -ErrorAction SilentlyContinue | Where-Object {
    $_.CommandLine -like "*pi4office*server*"
  } | Stop-Process -Force -ErrorAction SilentlyContinue

  # Remove shortcut
  if (Test-Path $SHORTCUT_PATH) {
    Remove-Item $SHORTCUT_PATH -Force
    Write-Step "Removed start menu shortcut"
  }

  # Remove manifest from wef
  $wefDirs = Get-OfficeWefDir
  foreach ($wef in $wefDirs) {
    $manifestPath = "$($wef.Path)\pi4office-manifest.xml"
    if (Test-Path $manifestPath) {
      Remove-Item $manifestPath -Force
      Write-Step "Removed manifest from $($wef.App) wef folder"
    }
  }

  # Remove app directory
  if (Test-Path $APP_DIR) {
    Remove-Item $APP_DIR -Recurse -Force
    Write-Step "Removed application directory: $APP_DIR"
  }

  Write-Host ""
  Write-Host "  Pi for Office has been uninstalled." -ForegroundColor Green
  Write-Host ""
  exit 0
}

function Invoke-Install {
  Write-Host ""
  Write-Host "  =============================================" -ForegroundColor Cyan
  Write-Host "    Pi for Office - Windows Installer" -ForegroundColor Cyan
  Write-Host "  =============================================" -ForegroundColor Cyan
  Write-Host ""

  # Check if running from bundled installer
  $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
  $isBundled = Test-Path "$scriptDir\dist" -PathType Container

  # Create app directory
  Write-Step "Installing to: $APP_DIR"
  New-Item -ItemType Directory -Path $APP_DIR -Force | Out-Null

  # Copy application files if bundled
  if ($isBundled) {
    Write-Step "Copying application files..."

    # Copy dist/
    if (-not (Test-Path "$APP_DIR\dist")) {
      Copy-Item -Path "$scriptDir\dist" -Destination "$APP_DIR\dist" -Recurse -Force
      Write-Check "Copied application files"
    }

    # Copy server
    Copy-Item -Path "$scriptDir\server.mjs" -Destination $SERVER_SCRIPT -Force
    Write-Check "Copied server script"

    # Copy nodejs if included
    if ((Test-Path "$scriptDir\nodejs") -and (-not (Test-Path $NODE_PORTABLE_DIR))) {
      Copy-Item -Path "$scriptDir\nodejs" -Destination $NODE_PORTABLE_DIR -Recurse -Force
      Write-Check "Copied Node.js runtime"
    }

    # Copy mkcert if included
    if ((Test-Path "$scriptDir\mkcert.exe") -and (-not (Test-Path $MKCERT_EXE))) {
      Copy-Item -Path "$scriptDir\mkcert.exe" -Destination $MKCERT_EXE -Force
      Write-Check "Copied mkcert"
    }
  }

  # Check for Node.js
  $nodeExe = $null
  if (Test-Path $NODE_EXE) {
    $nodeExe = $NODE_EXE
    Write-Check "Using bundled Node.js: $NODE_EXE"
  } else {
    $nodeExe = (Get-Command "node.exe" -ErrorAction SilentlyContinue)?.Source
    if (-not $nodeExe) {
      $nodeExe = (Get-Command "node" -ErrorAction SilentlyContinue)?.Source
    }
    if ($nodeExe) {
      Write-Check "Found Node.js: $nodeExe"
    }
  }

  if (-not $nodeExe) {
    Write-Warn "Node.js not found. Some features require Node.js."
    Write-Warn "Install Node.js from https://nodejs.org (v20+) or use the full installer."
  }

  # Generate TLS certificates
  Write-Step "Setting up HTTPS certificates..."
  New-Item -ItemType Directory -Path $CERT_DIR -Force | Out-Null

  if (-not (Test-Path "$CERT_DIR\key.pem") -or -not (Test-Path "$CERT_DIR\cert.pem")) {
    if (Test-Path $MKCERT_EXE) {
      # Use bundled mkcert
      Write-Step "Installing mkcert CA..."
      & $MKCERT_EXE -install 2>&1 | Out-Null
      if ($LASTEXITCODE -ne 0) {
        Write-Warn "mkcert CA install had warnings (may be OK)"
      }

      Write-Step "Generating localhost certificate..."
      Push-Location $CERT_DIR
      & $MKCERT_EXE -key-file key.pem -cert-file cert.pem localhost 127.0.0.1 ::1
      Pop-Location

      if ((Test-Path "$CERT_DIR\key.pem") -and (Test-Path "$CERT_DIR\cert.pem")) {
        Write-Check "Certificates generated successfully"
      } else {
        Write-Err "Certificate generation failed"
        Write-Warn "You can generate certs manually with: mkcert -install && mkcert localhost"
      }
    } else {
      Write-Warn "mkcert not available - attempting self-signed certificate..."

      # Fallback: generate self-signed cert using PowerShell
      try {
        $cert = New-SelfSignedCertificate -DnsName "localhost" `
          -CertStoreLocation "Cert:\CurrentUser\My" `
          -KeyUsage DigitalSignature, KeyEncipherment `
          -TextExtension "2.5.29.37={text}1.3.6.1.5.5.7.3.1" `
          -NotAfter (Get-Date).AddYears(5)

        # Export as PFX then convert to PEM
        $pfxPath = "$CERT_DIR\localhost.pfx"
        $pfxPassword = ConvertTo-SecureString -String "pi4office" -Force -AsPlainText
        Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pfxPassword | Out-Null

        # Use openssl if available
        $openssl = Get-Command "openssl.exe" -ErrorAction SilentlyContinue
        if ($openssl) {
          & openssl pkcs12 -in $pfxPath -nocerts -out "$CERT_DIR\key.pem" `
            -passin pass:pi4office -passout pass:pi4office -nodes 2>&1 | Out-Null
          & openssl pkcs12 -in $pfxPath -clcerts -nokeys -out "$CERT_DIR\cert.pem" `
            -passin pass:pi4office 2>&1 | Out-Null
          Write-Check "Self-signed certificates generated (requires trusting manually)"
        } else {
          Write-Warn "OpenSSL not available - certificates must be created manually"
          Write-Warn "Run: mkcert -install && mkcert localhost"
        }

        Remove-Item $cert.PSPath -Force -ErrorAction SilentlyContinue
      } catch {
        Write-Warn "Self-signed certificate generation failed: $_"
        Write-Warn "Install mkcert and re-run: winget install FiloSottile.mkcert"
      }
    }
  } else {
    Write-Check "Certificates already exist"
  }

  # Write manifest
  Write-Step "Writing manifest..."
  $manifestContent = $MANIFEST_CONTENT.Replace("3141", $Port.ToString())
  Set-Content -Path $MANIFEST_FILE -Value $manifestContent -Encoding UTF8

  # Copy manifest to Office wef folders
  Write-Step "Registering with Office..."
  $wefDirs = Get-OfficeWefDir
  foreach ($wef in $wefDirs) {
    $destManifest = "$($wef.Path)\pi4office-manifest.xml"
    Copy-Item -Path $MANIFEST_FILE -Destination $destManifest -Force
    Write-Check "Registered in $($wef.App) ($($wef.Path))"
  }

  if ($wefDirs.Count -eq 0) {
    Write-Warn "No Office wef folder found. Manually add the manifest:"
    Write-Warn "  Excel -> Insert -> My Add-ins -> Upload My Add-in"
    Write-Warn "  Select: $MANIFEST_FILE"
  }

  # Create start menu shortcut
  Write-Step "Creating start menu shortcut..."
  $WshShell = New-Object -ComObject WScript.Shell
  $Shortcut = $WshShell.CreateShortcut($SHORTCUT_PATH)
  if ($nodeExe) {
    $Shortcut.TargetPath = $nodeExe
    $Shortcut.Arguments = "`"$SERVER_SCRIPT`""
    $Shortcut.WorkingDirectory = $APP_DIR
  } else {
    $Shortcut.TargetPath = "powershell.exe"
    $Shortcut.Arguments = "-Command `"Write-Host 'Node.js is required to run Pi for Office server. Install from https://nodejs.org'; pause`""
  }
  $Shortcut.Description = "Start Pi for Office local server"
  $Shortcut.Save()
  Write-Check "Shortcut created: $SHORTCUT_PATH"

  Write-Host ""
  Write-Host "  =============================================" -ForegroundColor Green
  Write-Host "   Installation complete!" -ForegroundColor Green
  Write-Host "  =============================================" -ForegroundColor Green
  Write-Host ""
  Write-Host "  What's next:" -ForegroundColor Yellow
  Write-Host "  1. Start Pi for Office from the Start Menu" -ForegroundColor White
  Write-Host "     (or run: node `"$SERVER_SCRIPT`")" -ForegroundColor White
  Write-Host "  2. Open Excel, Word, or PowerPoint" -ForegroundColor White
  Write-Host "  3. Click 'Open Pi' in the Home ribbon" -ForegroundColor White
  Write-Host "  4. Connect a model provider and start chatting" -ForegroundColor White
  Write-Host ""
  Write-Host "  Local server URL: https://localhost:$Port" -ForegroundColor Cyan
  Write-Host "  Proxy URL: https://localhost:$Port/api-proxy/?url=<target>" -ForegroundColor Cyan
  Write-Host "  Settings -> Proxy: https://localhost:$Port/api-proxy/" -ForegroundColor Cyan
  Write-Host ""
}

# Main
if ($Uninstall) {
  Invoke-Uninstall
} else {
  Invoke-Install
}
