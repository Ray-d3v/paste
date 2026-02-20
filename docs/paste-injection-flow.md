# Paste Injection Flow (Current Working Behavior)

This document records the exact paste behavior that is currently working.
Code reference: `PasteWinUI/MainWindow.xaml.cs`.

## Trigger

- Overlay opens with `Ctrl+Alt+V`.
- Paste executes when `Enter` is pressed on a selected card.

Relevant methods:
- `PasteSelectedEntryAsync`
- `ResolvePasteTargetWindow`
- `HideOverlayForPaste`
- `InjectEntryIntoTarget`
- `TryInjectWithAttachedThreads`

## Target window selection

When opening overlay, the app stores the external window handle before focus moves to the overlay.

Paste target resolution order:
1. `_overlayOpenTargetWindow` (captured before opening overlay)
2. `_lastExternalForegroundWindow`
3. current foreground window (`GetForegroundWindow`) normalized to top-level (`GetAncestor(..., GA_ROOT)`)

Validation:
- target must be non-zero
- target must not be this app window
- `IsWindow(target)` must be true

## Enter paste sequence (exact order)

Inside `PasteSelectedEntryAsync`:
1. Guard against re-entry with `_isPastingSelection`.
2. Build clipboard payload from selected card via `TryApplyEntryToClipboardAsync`.
3. Immediately hide overlay for paste via `HideOverlayForPaste`:
   - stop animation
   - close preview/context/search (without focus restore)
   - mark overlay closed
   - move app window off-screen
   - call `_appWindow.Hide()` and `ShowWindow(..., SW_HIDE)`
4. Wait `80ms`.
5. Resolve/validate paste target; fallback to current foreground if needed.
6. `BringWindowToTop(target)`.
7. Call `InjectEntryIntoTarget(target)`.

## Clipboard write behavior

`TryApplyEntryToClipboardAsync`:
- retries up to 5 times (25ms delay between attempts)
- sets `_suppressClipboardUntilUtc` for 900ms to avoid self-capture in history
- writes:
  - image: `SetBitmap`
  - link: `SetWebLink` + `SetText`
  - text: `SetText`
- calls `Clipboard.Flush()`

## Input injection behavior (working path)

`InjectEntryIntoTarget` first tries `TryInjectWithAttachedThreads` (preferred path).

`TryInjectWithAttachedThreads`:
1. If target is minimized, restore it (`ShowWindow(..., SW_RESTORE)`).
2. Allow foreground switching (`AllowSetForegroundWindow(-1)`).
3. Attach input thread to target/foreground (`AttachThreadInput`).
4. Force activation:
   - `BringWindowToTop`
   - `SetForegroundWindow`
   - `SetActiveWindow`
   - if available, `SetFocus(hwndFocus)` from `GetGUIThreadInfo`
5. Attempt paste in this order:
   - `PostMessage(WM_PASTE)` to focused control
   - `SendInput` for `Ctrl+V`
   - `SendInput` for `Shift+Insert`
6. Detach thread input in `finally`.

If attached-thread path fails, fallback path still does:
- (optional) activate target
- `PostMessage(WM_PASTE)`
- `Ctrl+V`
- `Shift+Insert`

## Important note

`TryInjectTextEntryDirectly` / `SendUnicodeTextInput` exist in code but are **not part of the active success path** for normal paste behavior.
