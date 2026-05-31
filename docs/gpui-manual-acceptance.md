# GPUI Manual Acceptance Checklist

Date: 2026-05-31

Use this checklist after the automated release verifier passes. It covers the
Windows shell interactions that are intentionally not asserted through the smoke
scripts because notification-area layout and overflow behavior varies by user
profile and Windows 11 settings.

Automation note: Computer Use was attempted on 2026-05-31. It could start
`PasteWinUI.exe`, but the hidden GPUI overlay and Windows notification area were
not exposed as targetable windows, so this checklist remains a manual shell
acceptance step for now.

## Prerequisite

Run the full verifier first:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/verify-gpui-release.ps1
```

Expected result: all verifier steps pass.

You can also use the preparation helper, which runs the release verifier unless
skipped, launches the local GPUI EXE, probes the tray icon rectangle, and prints
the recording/final-gate commands:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/prepare-gpui-manual-acceptance.ps1
```

For a faster setup after a recent verifier pass:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/prepare-gpui-manual-acceptance.ps1 -SkipReleaseVerifier
```

## Tray Physical Clicks

1. Start the Rust GPUI EXE, or use the preparation helper above:

   ```powershell
   powershell -ExecutionPolicy Bypass -File tools/scripts/build-local-exe.ps1
   ```

2. Confirm `PasteWinUI.exe` remains running.

   Optional diagnostic probe:

   ```powershell
   powershell -ExecutionPolicy Bypass -File tools/scripts/probe-gpui-tray-rect.ps1 -UseExistingProcess -KeepRunning
   powershell -ExecutionPolicy Bypass -File tools/scripts/smoke-gpui-tray-click.ps1 -KeepRunning
   powershell -ExecutionPolicy Bypass -File tools/scripts/probe-gpui-tray-ui.ps1
   ```

   `probe-gpui-tray-rect.ps1` asks Windows Shell for the registered
   notification icon rectangle through `Shell_NotifyIconGetRect`. A successful
   result proves the app registered a tray icon and gives a concrete screen
   rectangle that can help locate the click target.

   `smoke-gpui-tray-click.ps1` attempts synthetic mouse left/right clicks at
   that rectangle. Treat a pass as useful supporting evidence, but do not treat
   a failure as an acceptance failure by itself: in the Codex desktop session on
   2026-05-31, `SendInput` did not reach the Windows notification area even
   though the Shell returned a valid icon rectangle.

   This probes Windows UI Automation and Win32 window enumeration for visible
   notification-area related elements. A matching result can help locate the
   icon, but a missing match does not fail acceptance because hidden overflow
   icons, and in some automation sessions the taskbar itself, are not exposed to
   automation.

3. Locate the `PasteWinUI` notification icon.

   If the icon is hidden in the Windows 11 overflow flyout, open the overflow
   flyout first. This is acceptable for the first GPUI release because Windows
   controls tray icon placement.

4. Left-click the tray icon.

   Expected: the Paste overlay appears.

5. Right-click the tray icon.

   Expected: a menu appears with `Show`, `Restart`, and `Exit`.

6. Click `Show`.

   Expected: the Paste overlay appears or remains visible.

7. Right-click the tray icon again and click `Restart`.

   Expected: the current `PasteWinUI.exe` process exits and a replacement
   `PasteWinUI.exe` process starts. This may not produce a visible UI change;
   verify by checking that the process ID changed.

   Helper:

   ```powershell
   powershell -ExecutionPolicy Bypass -File tools/scripts/check-gpui-restart-pid.ps1 -StartIfMissing -UpdateManualAcceptance -WaitSeconds 90
   ```

   Run the helper, then click `Restart` from the physical tray menu while it is
   waiting. If the icon is in the overflow flyout, open the overflow and choose
   `Restart` before the wait timeout expires. It prints the before/after process
   IDs and exits successfully only when a replacement process is observed. With
   `-UpdateManualAcceptance`, it marks `tray_menu_restart` as passed in the
   manual acceptance record. It also writes evidence to:

   ```text
   artifacts/manual-acceptance/gpui-restart-pid-check.json
   ```

8. Right-click the tray icon again and click `Exit`.

   Expected: `PasteWinUI.exe` exits and no replacement process remains.

## Overlay Visual Check

1. Focus any normal text target such as Notepad.
2. Press `Ctrl+Alt+V`.

Expected:

- The overlay appears near the bottom of the primary display.
- The overlay slides upward from below the bottom edge.
- Keyboard focus is inside the overlay.
- Pressing `Esc` hides the overlay immediately.

## Installer Physical Smoke

Automated coverage: `tools/scripts/verify-gpui-release.ps1` runs
`smoke-gpui-offline-installer.ps1`, which installs the generated offline
installer into `artifacts/installer-smoke`, verifies shortcuts, Run value,
uninstall registry metadata, installed EXE launch, and uninstall cleanup.

Use this physical check only for the default interactive installer path:

1. Run `deploy/installer/output/PasteWinUI-Setup-0.1.0-offline.exe`.
2. Accept or decline the desktop shortcut prompt.
3. Accept or decline auto-start.

Expected:

- The app installs to `%LOCALAPPDATA%\Programs\PasteWinUI`.
- The Start Menu shortcut is created.
- The uninstall entry is created under
  `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\PasteWinUI`.
- The app starts after installation.

## WinUI Fallback Decision

Decision for the first GPUI release candidate: keep the WinUI project and
`tools/scripts/build-local-winui-exe.ps1` as a comparison fallback. Do not delete
WinUI until the GPUI build has passed automated release verification and this
manual checklist on the target machine.

## Record Result

After completing the checklist, record the result:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/record-gpui-manual-acceptance.ps1
```

If the record already exists and only failed checks need to be retried:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/record-gpui-manual-acceptance.ps1 -RetryFailedOnly
```

The script writes:

```text
artifacts/manual-acceptance/gpui-manual-acceptance.json
```

Use `-AssumePassed` only when the checklist has already been completed and the
script is being used to create a machine-readable record afterward. The record
separates checks into `physical_shell`, `visual_manual`, and
`physical_installer` categories so the final gate can distinguish physical
Windows shell acceptance from automated smoke coverage.

Then run the final acceptance gate:

```powershell
powershell -ExecutionPolicy Bypass -File tools/scripts/verify-gpui-acceptance.ps1
```

Expected result: the release verifier passes and the manual acceptance record is
accepted.
