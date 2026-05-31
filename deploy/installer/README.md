# Installer Build

## Prerequisites
- Rust toolchain
- Inno Setup 6 (`ISCC.exe`)

## Build Steps
Run from repository root:

```powershell
.\tools\scripts\build-installer.ps1 -Version 0.1.0
```

This script does the following:
- builds the Rust GPUI app through `tools\scripts\build-local-exe.ps1`
- compiles `deploy/installer/PasteWinUI.GPUI.iss`
- outputs installer to `deploy/installer/output/`

## Useful Options
- Skip build and only rebuild installer:

```powershell
.\tools\scripts\build-installer.ps1 -Version 0.1.0 -SkipBuild
```

## WinUI Fallback
The retired WinUI local publish path is kept for comparison while the Rust
replacement is validated:

```powershell
.\tools\scripts\build-local-winui-exe.ps1 -NoLaunch
```

## Offline Alternative (No Inno Setup)
If Inno Setup is not installed, use Windows built-in IExpress:

```powershell
.\tools\scripts\build-offline-installer.ps1 -Version 0.1.0
```

Output:
- `deploy/installer/output/PasteWinUI-Setup-<version>-offline.exe`

Smoke-test the generated offline installer without touching the default install
location:

```powershell
.\tools\scripts\smoke-gpui-offline-installer.ps1
```

The smoke test installs into `artifacts\installer-smoke`, verifies shortcuts,
the Run value, uninstall registry metadata, installed app launch, and generated
uninstall cleanup.

## Notes
- Installer is per-user (`%LocalAppData%\Programs\PasteWinUI`) and does not require admin rights.
- Installer can optionally create:
  - desktop shortcut
  - auto-start entry (`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`)
- Installer also registers a simple uninstall entry:
  - `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\PasteWinUI`
  - Start Menu shortcut: `Uninstall PasteWinUI`
