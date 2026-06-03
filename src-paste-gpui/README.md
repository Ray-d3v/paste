# Paste GPUI

Rust + GPUI replacement workspace for the existing WinUI Paste app.

The first milestone is a Windows feasibility gate: build and launch the
minimal GPUI app before deleting or replacing the existing WinUI project.

## Crates

- `paste-core`: UI-independent clipboard history logic.
- `paste-windows-platform`: Win32 integration boundary.
- `paste-gpui-app`: GPUI overlay shell. Release binary name is `PasteWinUI.exe`.

## Overlay Rendering

The history view uses GPUI `uniform_list` with a tracked scroll handle, so row
elements are generated for the visible range instead of rebuilding every history
row on each render. Selection changes request `ScrollStrategy::Nearest` to keep
keyboard navigation visible without forcing a full list render.

The overlay window is created as a non-resizable GPUI floating window centered
near the bottom of the primary display. Because GPUI's Windows hide/show path is
not sufficient for this overlay workflow yet, the app uses a small Win32
positioning wrapper: opening places the native HWND just below the screen edge,
shows it, then moves it upward over a short eased animation. Closing still hides
the HWND immediately to preserve fast keyboard-first paste behavior.

`Ctrl+K` opens a keyboard command palette. While it is open, Up/Down select a
command, Enter runs the selected command, and Esc closes only the palette.

## Local Build

From the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/build-local-exe.ps1 -NoLaunch
```

This builds the Rust app and copies `PasteWinUI.exe` to `artifacts/local-gpui/`.

Full local release verification is available with:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/verify-gpui-release.ps1
```

This runs the release build, Rust workspace tests, and all GPUI smoke checks sequentially.
It also builds the offline IExpress installer from the generated Rust EXE. The
Inno Setup installer path remains available through `tools/scripts/build-installer.ps1`
when `ISCC.exe` is installed.

Latest full verification on 2026-05-31:

- `build-local-exe`: passed
- `cargo-test-workspace`: passed
- `smoke-hotkey`: passed
- `smoke-slide-up`: passed
- `smoke-command-palette`: passed
- `smoke-text-paste`: passed
- `smoke-link-paste`: passed
- `smoke-image-paste`: passed
- `smoke-file-paste`: passed
- `smoke-tray-show-exit`: passed
- `smoke-tray-rect`: passed
- `smoke-tray-notification`: passed
- `smoke-tray-restart`: passed
- `build-offline-installer`: passed
- `smoke-offline-installer`: passed
- Offline installer output: `deploy/installer/output/PasteWinUI-Setup-0.1.0-offline.exe`

Requirement-by-requirement status is tracked in
`docs/gpui-migration-audit.md`. Manual shell checks that cannot be made stable
across all Windows notification-area layouts are tracked in
`docs/gpui-manual-acceptance.md`. The final combined gate is
`tools/scripts/verify-gpui-acceptance.ps1`.

To prepare the final physical Windows shell checks, run:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/prepare-gpui-manual-acceptance.ps1
```

This launches the local GPUI EXE, verifies the tray icon rectangle through
`Shell_NotifyIconGetRect`, and prints the manual recording and final gate
commands.

For the physical `Restart` tray-menu check, use:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/check-gpui-restart-pid.ps1 -StartIfMissing
```

Then click `Restart` from the physical tray menu while the helper is waiting.

The legacy WinUI comparison build is still available:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/build-local-winui-exe.ps1 -NoLaunch
```

Release builds of GPUI on Windows need `fxc.exe` from the Windows SDK for shader
compilation. The build script automatically adds the latest installed Windows
SDK `x64` FXC directory to `PATH` when needed.

## Release Size Snapshot

- Debug GPUI EXE: 31.76 MB
- Release GPUI EXE with `strip`, thin LTO, `opt-level = "s"`: 6.88 MB

## Local Performance Snapshot

Measured with:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/measure-local-exe.ps1 -SampleDelayMs 1500
```

- Rust GPUI release EXE: 6.88 MB
- Rust GPUI deployment footprint: 6.88 MB
- Rust GPUI ready signal: hidden Win32 message window observed at ~361 ms
- Rust GPUI working set after ~1.5s sample: 61.09 MB
- Rust GPUI peak working set after ~1.5s sample: 62.49 MB
- Legacy WinUI EXE stub: 0.26 MB
- Legacy WinUI deployment footprint: 160.07 MB
- Legacy WinUI ready signal: main window observed at ~1396 ms
- Legacy WinUI working set after ~1.5s sample: 127.58 MB
- Legacy WinUI peak working set after ~1.5s sample: 127.58 MB

Comparison from this local sample:

- Rust GPUI deployment footprint is ~153.19 MB smaller.
- Rust GPUI ready signal is ~1035 ms faster.
- Rust GPUI working set is ~66.49 MB lower.

## Hotkey Smoke Check

Measured with:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-hotkey.ps1
powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-slide.ps1
powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-command-palette.ps1
```

Result:

- Rust GPUI EXE launched successfully
- `Ctrl+Alt+V` was sent through the script fallback path
- Process stayed alive after hotkey dispatch
- A main window handle was observed
- The overlay HWND was hidden before the hotkey and visible after the hotkey
- The slide smoke sampled the overlay HWND after hotkey dispatch and verified upward travel before settling near the highest sampled position
- Latest local slide sample moved the overlay top from 1064px to 628px and settled at 628px during the aggregate smoke run
- The command palette smoke verifies `Ctrl+K`, Esc closing the palette without hiding the overlay, and Enter executing a palette command

Limitations:

- The script does not inspect rendered pixels or prove the overlay is visually in the expected position.
- The slide smoke verifies native window movement, not rendered frame contents.
- The primary `SendInput` inline PowerShell path returned 0 events in this environment, so the script fell back to `keybd_event`.

## Tray Smoke Check

Measured with:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-tray.ps1
powershell -ExecutionPolicy Bypass -File tools/scripts/probe-gpui-tray-rect.ps1
powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-tray-notification.ps1
powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-tray-restart.ps1
```

Result:

- Rust GPUI EXE launched successfully
- The hidden Win32 message window was found
- The tray Show command made the overlay HWND visible
- `Shell_NotifyIconGetRect` returned the registered tray icon rectangle
- The `NOTIFYICON_VERSION_4` `NIN_SELECT` notification showed the overlay and `WM_CONTEXTMENU` opened a Win32 popup menu
- The tray Exit command shut down `PasteWinUI.exe` with exit code 0
- The tray Restart command shut down the original process with exit code 0 and launched a replacement `PasteWinUI.exe`

Limitations:

- This verifies the app's tray command handling through its Win32 message loop, verifies that Windows Shell can return the tray icon rectangle, and verifies the notification messages used by physical tray selection/context-menu activation. It does not physically click the notification area icon, which is shell-layout dependent on Windows 11.
- `tools/scripts/smoke-gpui-tray-click.ps1` is available as a best-effort synthetic mouse-input diagnostic. In the Codex desktop session on 2026-05-31, `SendInput` did not reach the Windows notification area even though `Shell_NotifyIconGetRect` returned a valid icon rectangle.
- The tray smoke scripts manage the singleton `PasteWinUI.exe`; run them sequentially, not in parallel.

## Text Paste Smoke Check

Measured with:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-paste-text.ps1 -OverlayDelayMs 1800 -PasteDelayMs 1800
```

Result:

- Rust GPUI EXE launched successfully
- A temporary external WinForms text target launched successfully
- The script copied a unique text value, dispatched `Ctrl+Alt+V`, pressed `Enter`, and verified the target received the exact text

Limitations:

- This proves text capture and paste injection into a simple external edit control; it does not yet cover image paste, link metadata preview, or elevated/UIPI-protected targets.
- GPUI's Windows `App::activate`/`App::hide` path is currently insufficient for hidden overlay toggling, so the app uses a Win32 title/PID lookup to call native `ShowWindow` for overlay visibility.

## Link Paste Smoke Check

Measured with:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-paste-link.ps1
```

Result:

- Rust GPUI EXE launched successfully
- A temporary external WinForms text target launched successfully
- The script copied a unique URL, dispatched `Ctrl+Alt+V`, pressed `Enter`, and verified the target received the exact URL
- The persisted latest history entry was verified as `kind = Link` with `link_url` equal to the copied URL

Limitations:

- This proves URL classification, persistence, and paste behavior; network page title fetching is not implemented yet, so `link_title` remains optional/empty.

## Image Paste Smoke Check

Measured with:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-paste-image.ps1
```

Result:

- Rust GPUI EXE launched successfully
- A temporary external WinForms image target launched successfully
- The script placed a 16x12 bitmap on the clipboard, dispatched `Ctrl+Alt+V`, pressed `Enter`, and verified the target received a matching CF_DIB image
- WinForms-provided 32bpp `BI_BITFIELDS` DIBs are accepted by the core PNG conversion path

Limitations:

- This proves CF_BITMAP/CF_DIB image capture and paste into a simple target; it does not yet cover large images, alpha-sensitive image workflows, or application-specific rich image formats.
- Text, hotkey, and image smoke scripts each manage the singleton `PasteWinUI.exe`; run them sequentially, not in parallel.

## File Paste Smoke Check

Measured with:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-paste-file.ps1
```

Result:

- Rust GPUI EXE launched successfully
- A temporary external WinForms file-drop target launched successfully
- The script created two temporary files, placed them on the clipboard as a file drop list, dispatched `Ctrl+Alt+V`, pressed `Enter`, and verified the target received a matching `CF_HDROP` file list
- The persisted matching history entry was verified as `kind = File` with the expected `file_paths`

Limitations:

- This proves existing file paths are captured and pasted as `CF_HDROP`; virtual file transfer formats such as `CFSTR_FILEDESCRIPTOR`/`CFSTR_FILECONTENTS` are not implemented.
- Text, hotkey, image, and file smoke scripts each manage the singleton `PasteWinUI.exe`; run them sequentially, not in parallel.

## Offline Installer Smoke Check

Measured with:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-offline-installer.ps1
```

Result:

- The generated IExpress offline installer exited with code 0.
- The package installed to a temporary smoke directory under `artifacts/installer-smoke`.
- `PasteWinUI.exe`, Start Menu shortcut, uninstall shortcut, desktop shortcut override, Run value, uninstall script, and uninstall registry key were created.
- The installed EXE started successfully from the smoke install directory.
- Running the generated uninstall script removed the install directory, Start Menu directory, desktop shortcut, Run value, and uninstall key.
