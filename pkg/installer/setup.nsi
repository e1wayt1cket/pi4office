; Pi for Office Windows Installer
; NSIS (Nullsoft Scriptable Install System) script
;
; Build: makensis setup.nsi
; Requires NSIS 3.x: winget install NSIS.NSIS

Unicode true
!include "MUI2.nsh"
!include "FileFunc.nsh"

; --- Installer Configuration ---
Name "Pi for Office"
OutFile "..\pi4office-setup.exe"
InstallDir "$LOCALAPPDATA\pi4office"
RequestExecutionLevel user
SetCompressor /SOLID lzma

!define APP_NAME "Pi for Office"
!define APP_DIR "pi4office"
!define SERVER_SCRIPT "server.mjs"
!define PORT "3141"

; --- Modern UI ---
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

; --- Installer Section ---
Section "Install"

  SetOutPath "$INSTDIR"

  ; Copy all application files
  File /r "dist\"
  File "server.mjs"
  File "mkcert.exe"

  ; Copy Node.js runtime
  SetOutPath "$INSTDIR\nodejs"
  File /r "nodejs\*.*"

  SetOutPath "$INSTDIR"

  ; Generate certificates
  DetailPrint "Setting up HTTPS certificates..."
  nsExec::ExecToLog '"$INSTDIR\mkcert.exe" -install'
  nsExec::ExecToLog '"$INSTDIR\mkcert.exe" -key-file "$INSTDIR\certs\key.pem" -cert-file "$INSTDIR\certs\cert.pem" localhost 127.0.0.1 ::1'

  ; Write manifest
  DetailPrint "Writing manifest..."
  FileOpen $0 "$INSTDIR\manifest.xml" w
  FileWrite $0 '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>$\r$\n'
  FileWrite $0 '<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"$\r$\n'
  FileWrite $0 '  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"$\r$\n'
  FileWrite $0 '  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"$\r$\n'
  FileWrite $0 '  xsi:type="TaskPaneApp">$\r$\n'
  FileWrite $0 '  <Id>a1b2c3d4-e5f6-7890-abcd-ef1234567890</Id>$\r$\n'
  FileWrite $0 '  <Version>1.0.0.0</Version>$\r$\n'
  FileWrite $0 '  <ProviderName>Pi for Office</ProviderName>$\r$\n'
  FileWrite $0 '  <DefaultLocale>en-US</DefaultLocale>$\r$\n'
  FileWrite $0 '  <DisplayName DefaultValue="Pi for Office" />$\r$\n'
  FileWrite $0 '  <Description DefaultValue="Open-source, multi-model AI assistant for Microsoft Office." />$\r$\n'
  FileWrite $0 '  <IconUrl DefaultValue="https://localhost:${PORT}/assets/icon-32.png" />$\r$\n'
  FileWrite $0 '  <HighResolutionIconUrl DefaultValue="https://localhost:${PORT}/assets/icon-80.png" />$\r$\n'
  FileWrite $0 '  <SupportUrl DefaultValue="https://github.com/tmustier/pi4office" />$\r$\n'
  FileWrite $0 '  <AppDomains>$\r$\n'
  FileWrite $0 '    <AppDomain>https://localhost</AppDomain>$\r$\n'
  FileWrite $0 '  </AppDomains>$\r$\n'
  FileWrite $0 '  <Hosts>$\r$\n'
  FileWrite $0 '    <Host Name="Workbook" />$\r$\n'
  FileWrite $0 '    <Host Name="Document" />$\r$\n'
  FileWrite $0 '    <Host Name="Presentation" />$\r$\n'
  FileWrite $0 '  </Hosts>$\r$\n'
  FileWrite $0 '  <DefaultSettings>$\r$\n'
  FileWrite $0 '    <SourceLocation DefaultValue="https://localhost:${PORT}/src/taskpane.html" />$\r$\n'
  FileWrite $0 '    <RequestedWidth>400</RequestedWidth>$\r$\n'
  FileWrite $0 '    <RequestedHeight>600</RequestedHeight>$\r$\n'
  FileWrite $0 '  </DefaultSettings>$\r$\n'
  FileWrite $0 '  <Permissions>ReadWriteDocument</Permissions>$\r$\n'
  FileWrite $0 '</OfficeApp>$\r$\n'
  FileClose $0

  ; Copy manifest to Office wef folders
  DetailPrint "Registering with Office..."
  CreateDirectory "$LOCALAPPDATA\Microsoft\Office\WEF"
  CopyFiles "$INSTDIR\manifest.xml" "$LOCALAPPDATA\Microsoft\Office\WEF\pi4office-manifest.xml"

  ; Copy to Microsoft Store Office path if exists
  ${If} ${FileExists} "$LOCALAPPDATA\Packages\Microsoft.Office.Desktop_*\LocalCache\Content\Microsoft\WEF\*.*"
    FindFirst $0 $1 "$LOCALAPPDATA\Packages\Microsoft.Office.Desktop_*\LocalCache\Content\Microsoft\WEF"
    ${If} $1 != ""
      CopyFiles "$INSTDIR\manifest.xml" "$1\pi4office-manifest.xml"
    ${EndIf}
    FindClose $0
  ${EndIf}

  ; Create start menu shortcut
  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortCut "$SMPROGRAMS\${APP_NAME}\Pi for Office Server.lnk" \
    "$INSTDIR\nodejs\node.exe" \
    '"$INSTDIR\${SERVER_SCRIPT}"' \
    "$INSTDIR\dist\assets\icon-32.png" 0
  CreateShortCut "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk" \
    "$INSTDIR\uninstall.exe"

  ; Create desktop shortcut
  CreateShortCut "$DESKTOP\Pi for Office.lnk" \
    "$INSTDIR\nodejs\node.exe" \
    '"$INSTDIR\${SERVER_SCRIPT}"' \
    "$INSTDIR\dist\assets\icon-32.png" 0

  ; Write uninstaller
  WriteUninstaller "$INSTDIR\uninstall.exe"

  ; Registry for Add/Remove Programs
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" \
    "DisplayName" "${APP_NAME}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" \
    "UninstallString" "$\"$INSTDIR\uninstall.exe$\""
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" \
    "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" \
    "DisplayIcon" "$INSTDIR\dist\assets\icon-32.png"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" \
    "Publisher" "Pi for Office"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" \
    "URLInfoAbout" "https://github.com/tmustier/pi4office"
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" \
    "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" \
    "NoRepair" 1

SectionEnd

; --- Uninstaller Section ---
Section "Uninstall"

  ; Stop running server
  nsExec::ExecToLog 'taskkill /f /im node.exe /fi "IMAGENAME eq node.exe"'

  ; Remove shortcuts
  Delete "$SMPROGRAMS\${APP_NAME}\Pi for Office Server.lnk"
  Delete "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk"
  RMDir "$SMPROGRAMS\${APP_NAME}"
  Delete "$DESKTOP\Pi for Office.lnk"

  ; Remove manifest from Office
  Delete "$LOCALAPPDATA\Microsoft\Office\WEF\pi4office-manifest.xml"
  FindFirst $0 $1 "$LOCALAPPDATA\Packages\Microsoft.Office.Desktop_*\LocalCache\Content\Microsoft\WEF"
  ${If} $1 != ""
    Delete "$1\pi4office-manifest.xml"
  ${EndIf}
  FindClose $0

  ; Remove application files
  RMDir /r "$INSTDIR"

  ; Remove registry
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"

SectionEnd
