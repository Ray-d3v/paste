# GPUI Migration Audit

Date: 2026-05-31

This document tracks the current Rust + GPUI replacement against the requested
Paste migration scope. It is intentionally evidence-oriented: an item is marked
verified only when a file or command proves the behavior.

## Current Verdict

Status: release-candidate, not final-complete until manual shell acceptance is
run on the target Windows profile.

The Rust + GPUI implementation builds, runs, passes unit tests, passes
integration smokes for hotkey, slide-up, text paste, link paste, image paste,
tray command handling, tray icon rectangle discovery, tray notification handling,
restart, and offline installer generation. The offline installer smoke also
verifies temporary install, launch, uninstall metadata, and cleanup. The
remaining release gate is:

- Physical notification-area click testing is not automated; tray behavior is
  verified through the app's Win32 message window and command handler. Manual
  acceptance steps are in `docs/gpui-manual-acceptance.md`. Computer Use was
  attempted on 2026-05-31; it could launch `PasteWinUI.exe`, but the hidden GPUI
  overlay and Windows notification area were not exposed as targetable windows,
  so no reliable physical tray click could be performed from automation. A later
  Computer Use retry on 2026-05-31 connected successfully and listed normal app
  windows, but `list_windows()` still returned no targetable taskbar,
  notification-area, overflow, tray, Shell, Paste, or GPUI window.
- First GPUI release decision: keep the WinUI project and
  `tools/scripts/build-local-winui-exe.ps1` as comparison fallback. Do not delete
  WinUI during the first GPUI release candidate.
- File clipboard items and advanced icon extraction remain out of first-version
  scope by plan.

## Verification Commands

Latest full command:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/verify-gpui-release.ps1
```

Latest full result on 2026-05-31:

- `build-local-exe`: passed
- `cargo-test-workspace`: passed
- `smoke-hotkey`: passed
- `smoke-slide-up`: passed
- `smoke-command-palette`: passed
- `smoke-text-paste`: passed
- `smoke-link-paste`: passed
- `smoke-image-paste`: passed
- `smoke-tray-show-exit`: passed
- `smoke-tray-rect`: passed
- `smoke-tray-notification`: passed
- `smoke-tray-restart`: passed
- `build-offline-installer`: passed
- `smoke-offline-installer`: passed

Generated artifacts:

- `artifacts/local-gpui/PasteWinUI.exe`
- `deploy/installer/output/PasteWinUI-Setup-0.1.0-offline.exe`

Manual acceptance helper:

- `tools/scripts/probe-gpui-tray-ui.ps1` enumerates visible UI Automation
  elements and Win32 windows related to the taskbar/notification area, then
  reports entries matching `Paste|PasteWinUI|GPUI`. It is diagnostic only;
  hidden overflow icons can be absent from the UI Automation tree.
  On 2026-05-31 in this Codex session, the probe returned zero UIA and zero
  Win32 taskbar/notification-area related items even while `PasteWinUI.exe` was
  running, confirming that physical tray acceptance must be performed manually
  on the interactive desktop.
- `tools/scripts/probe-gpui-tray-rect.ps1` asks Windows Shell for the
  registered notification icon rectangle through `Shell_NotifyIconGetRect`.
  It is now included in the release verifier. On 2026-05-31 it returned
  `HRESULT = 0x00000000` with a 32x48 icon rectangle.
- `tools/scripts/prepare-gpui-manual-acceptance.ps1` is the operator entrypoint
  for final physical checks. It runs the release verifier unless skipped,
  launches the GPUI EXE, probes the tray icon rectangle, and prints the
  recording/final-gate commands. A `-SkipReleaseVerifier` dry run on 2026-05-31
  found the running `PasteWinUI.exe` process and a valid 32x48 tray rectangle.
- `tools/scripts/smoke-gpui-tray-notification.ps1` posts the notification
  events produced by `NOTIFYICON_VERSION_4`: `NIN_SELECT` must show the overlay,
  and `WM_CONTEXTMENU` must open a visible `#32768` popup menu. Latest run
  passed with one visible popup menu.
- `tools/scripts/smoke-gpui-tray-click.ps1` attempts a synthetic mouse click at
  the `Shell_NotifyIconGetRect` center. In the Codex desktop session on
  2026-05-31, `SendInput` did not reach the notification area, so this remains a
  best-effort diagnostic and is not part of the release gate.
- `tools/scripts/record-gpui-manual-acceptance.ps1` records completed manual
  checks to `artifacts/manual-acceptance/gpui-manual-acceptance.json`. The goal
  can be considered fully accepted only after that record exists with
  `all_passed = true`. Checks are categorized as `physical_shell`,
  `visual_manual`, or `physical_installer`.
- `tools/scripts/verify-gpui-acceptance.ps1` is the final gate. It runs the
  release verifier unless skipped, then requires the manual acceptance record to
  exist with `all_passed = true` and all required check IDs present.
  Its missing-check error includes the actual check IDs found in the record.

## Requirement Trace

| Requirement | Status | Evidence |
| --- | --- | --- |
| Keep WinUI until GPUI gate passes | Verified | Existing WinUI project remains; Rust work is isolated under `src-paste-gpui/`. |
| Windows GPUI feasibility gate | Verified | `tools/scripts/verify-gpui-release.ps1` runs `build-local-exe`; latest full run passed. |
| New Rust workspace | Verified | `src-paste-gpui/Cargo.toml` with crates `paste-core`, `paste-windows-platform`, `paste-gpui-app`. |
| Replacement-friendly binary name | Verified | `paste-gpui-app` builds `PasteWinUI.exe`; local output is `artifacts/local-gpui/PasteWinUI.exe`. |
| Core/UI/platform split | Verified | `paste-core` holds models and pure logic; `paste-windows-platform` holds Win32 integration; `paste-gpui-app` holds GPUI UI and state. |
| New JSON history store | Verified | `paste-core::default_history_path` targets `%LOCALAPPDATA%\PasteGPUI\history.json`; app loads/saves via `load_history_document` and `save_history_document`. |
| No legacy history migration | Verified | No migration path from WinUI data exists in the Rust app; new store is independent. |
| Entry fields | Verified | `ClipboardEntry` includes `kind`, `content`, `copied_at`, `source_app`, `image_png_bytes`, `link_url`, `link_title`, `pinned_group_id`, `is_deleted`, `deleted_at`; Rust app also stores `image_dib_bytes` for Windows image paste fidelity. |
| Debounced persistence | Verified | `paste-gpui-app` uses `HISTORY_SAVE_DEBOUNCE` and `schedule_history_save`. |
| History cap and trash retention | Verified | `paste-core::RetentionPolicy` and `enforce_retention`; unit test `retention_drops_old_trash_and_caps_active_history`. |
| Duplicate detection | Verified | `paste-core::is_likely_duplicate`; unit test `duplicate_link_ignores_fragment_and_trailing_slash`. |
| Search matching | Verified | `paste-core::entry_matches_query` and `visible_entries`; unit test `entry_matches_link_metadata`. UI uses `Ctrl+F` and `search_query`. |
| Pinned group create/rename/delete | Verified | `create_group`, `rename_group`, `delete_group`; unit tests cover create and delete behavior. |
| Initial pinned groups | Verified | `seeded_groups` defines Quick, Work, Idea; UI binds `Ctrl+1`, `Ctrl+2`, `Ctrl+3`, `Ctrl+0`. |
| Trash restore/retention | Verified | UI `DeleteSelected` and `RestoreSelected`; retention unit test covers trash expiry. |
| Paste risk detection | Verified | `paste_core::paste_risk_reason`; unit test `paste_risk_detects_multiline_text`. |
| Text clipboard support | Verified | `read_text_from_clipboard`, `write_text_to_clipboard`; `smoke-gpui-paste-text.ps1` passed. |
| Link clipboard support | Verified | URL classification in `entry_from_clipboard_text`; `smoke-gpui-paste-link.ps1` verifies paste and persisted `kind = Link`. |
| Image clipboard support | Verified | CF_DIB/CF_BITMAP read/write in `paste-windows-platform`; `smoke-gpui-paste-image.ps1` passed. |
| File clipboard support | Out of scope | Plan explicitly defers file items. |
| Advanced icon extraction | Out of scope | Plan explicitly defers advanced icon extraction. |
| Global hotkey `Ctrl+Alt+V` | Verified | `RegisterHotKey` in `paste-windows-platform`; `smoke-gpui-hotkey.ps1` passed. |
| Clipboard monitoring | Verified | `AddClipboardFormatListener` in message loop; smokes capture text/link/image after clipboard changes. |
| Clipboard write-back | Verified | Text and image write functions in platform layer; text/link/image paste smokes passed. |
| Target window resolution | Verified | App stores `overlay_open_target` before opening; platform uses `GetForegroundWindow`, `GetAncestor`, `SetForegroundWindow`. Text paste smoke verifies paste returns to external target. |
| Paste injection order | Verified | `paste_into_window` uses `PostMessage(WM_PASTE)`, then `SendInput Ctrl+V`, then `SendInput Shift+Insert`, matching `docs/paste-injection-flow.md`. |
| Tray resident menu | Partially verified | `Shell_NotifyIconW`, popup menu, show/restart/exit command handling implemented. `smoke-gpui-tray.ps1` and `smoke-gpui-tray-restart.ps1` passed through message-window command injection. `probe-gpui-tray-rect.ps1` verifies `Shell_NotifyIconGetRect` returns the registered tray icon rectangle. `smoke-gpui-tray-notification.ps1` verifies `NIN_SELECT` shows the overlay and `WM_CONTEXTMENU` opens the popup menu. Computer Use launch/inspection and synthetic `SendInput` clicks on 2026-05-31 did not reach the notification area, so physical shell clicks remain manual. `probe-gpui-tray-ui.ps1` and `smoke-gpui-tray-click.ps1` are available as additional diagnostics before manual acceptance. |
| Bottom floating overlay | Verified | `WindowKind::Floating`, `bottom_overlay_bounds`, native Win32 show/hide; hotkey smoke verifies hidden-before/open-after. |
| Slide-up on show | Verified | `animate_native_overlay_window` and `SetWindowPos` wrapper; `smoke-gpui-slide.ps1` passed. |
| Fast hide on close/paste | Verified by implementation | `close_overlay`, toggle close, and paste path call native hide immediately. No separate smoke asserts close timing. |
| Keyboard-first commands | Verified by implementation | Key bindings cover Enter, Esc, arrows, Ctrl+F, Ctrl+T, Ctrl+K, Ctrl+Shift+V, Alt+Enter, Delete, Ctrl+R, group pins. Paste smokes verify Enter path. |
| Command palette | Verified | `Ctrl+K` toggles a keyboard command palette. While open, Up/Down moves command selection, Enter runs the selected command, and Esc closes the palette without hiding the overlay. `smoke-gpui-command-palette.ps1` covers the keyboard behavior. |
| Virtualized/limited history rendering | Verified | GPUI `uniform_list` with `UniformListScrollHandle`; documented in `src-paste-gpui/README.md`. |
| Build script switched to Rust output | Verified | `tools/scripts/build-local-exe.ps1` builds Rust GPUI and copies to `artifacts/local-gpui/PasteWinUI.exe`; full verifier passed. |
| Installer script uses Rust output | Verified | `tools/scripts/build-offline-installer.ps1` packages `artifacts/local-gpui`; full verifier generated offline installer. |
| Offline installer installs and uninstalls | Verified | `smoke-gpui-offline-installer.ps1` installs the generated IExpress package to `artifacts/installer-smoke`, verifies EXE, shortcuts, Run value, uninstall key, app launch, then runs uninstall and verifies cleanup. |
| Release performance comparison | Verified | `tools/scripts/measure-local-exe.ps1` and `src-paste-gpui/README.md` record Rust vs WinUI size, ready time, and memory. |

## Residual Risks

- GPUI is still pre-1.0 and pulled from the Zed dependency path, so API churn is a real maintenance risk.
- The shell notification area is layout-dependent on Windows 11; current automated tray smokes validate command handling, Shell icon rectangle discovery, and the notification messages produced by real tray selection/context-menu activation, but not a real mouse click through the taskbar overflow UI. Computer Use connected successfully but did not expose taskbar/notification-area windows, and synthetic `SendInput` clicks could not target the notification area in this session. Manual acceptance is documented in `docs/gpui-manual-acceptance.md`.
- UIPI/elevated target paste behavior is not covered by current smokes.
- Link title fetching is not implemented; link entries store `link_url` and may leave `link_title` empty.
- The offline installer is an IExpress package. It is useful for local verification but less polished than a production installer built with Inno Setup or MSI tooling.

## Next Release Decisions

1. Run `docs/gpui-manual-acceptance.md` on the target Windows profile and record
   the result with `tools/scripts/record-gpui-manual-acceptance.ps1`.
2. Run `tools/scripts/verify-gpui-acceptance.ps1` and require it to pass before
   declaring the migration accepted.
3. After acceptance, decide whether to remove WinUI in a later cleanup change or keep it for one more comparison cycle.
4. If publishing externally, decide whether IExpress is sufficient or require Inno Setup installation on the build machine.
