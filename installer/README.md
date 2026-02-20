# Installer Build

## Prerequisites
- .NET SDK 8.x
- Inno Setup 6 (`ISCC.exe`)

## Build Steps
Run from repository root:

```powershell
.\scripts\build-installer.ps1 -Version 0.1.0
```

This script does the following:
- `dotnet publish` for `PasteWinUI` (`win-x64`, self-contained)
- compiles `installer/PasteWinUI.iss`
- outputs installer to `installer/output/`

## Useful Options
- Skip publish and only rebuild installer:

```powershell
.\scripts\build-installer.ps1 -Version 0.1.0 -SkipPublish
```

## Offline Alternative (No Inno Setup)
If Inno Setup is not installed, use Windows built-in IExpress:

```powershell
.\scripts\build-offline-installer.ps1 -Version 0.1.0
```

Output:
- `installer/output/PasteWinUI-Setup-<version>-offline.exe`

## Notes
- Installer is per-user (`%LocalAppData%\Programs\PasteWinUI`) and does not require admin rights.
- Installer can optionally create:
  - desktop shortcut
  - auto-start entry (`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`)
- Installer also registers a simple uninstall entry:
  - `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\PasteWinUI`
  - Start Menu shortcut: `Uninstall PasteWinUI`
