use gpui::prelude::*;
use gpui::{
    actions, bounds, div, point, px, rgb, size, uniform_list, App, Bounds, Context, FocusHandle,
    Focusable, IntoElement, KeyBinding, Pixels, Render, ScrollStrategy, UniformListScrollHandle,
    Window, WindowBounds, WindowKind, WindowOptions,
};
use gpui_platform::application;
use paste_core::{
    create_document, default_history_path, dib_to_png_bytes, enforce_retention,
    is_likely_duplicate, load_history_document, save_history_document, seeded_groups,
    ClipboardEntry, ClipboardHistoryDocument, RetentionPolicy, IDEA_GROUP_ID, QUICK_GROUP_ID,
    WORK_GROUP_ID,
};
use paste_windows_platform::{
    spawn_platform_event_loop, PlatformEvent, PlatformEventLoop, PlatformIntegration, WindowHandle,
    WindowsPlatformIntegration,
};
use std::{
    env, io,
    path::PathBuf,
    process::Command,
    sync::mpsc::{Receiver, TryRecvError},
    time::Duration,
};
use time::OffsetDateTime;
use url::Url;

const HISTORY_SAVE_DEBOUNCE: Duration = Duration::from_millis(300);
const OVERLAY_WINDOW_TITLE: &str = "Paste GPUI Overlay";
const OVERLAY_PANEL_WIDTH: f32 = 960.0;
const OVERLAY_PANEL_HEIGHT: f32 = 420.0;
const OVERLAY_BOTTOM_MARGIN: f32 = 32.0;
const OVERLAY_ANIMATION_STEPS: u64 = 8;
const OVERLAY_ANIMATION_FRAME_MS: u64 = 14;

actions!(
    paste_overlay,
    [
        MoveUp,
        MoveDown,
        PasteSelected,
        CloseOverlay,
        ToggleTrash,
        TogglePlainText,
        ToggleCommandPalette,
        ToggleSearch,
        OneShotPlainPaste,
        DeleteSelected,
        RestoreSelected,
        PinQuick,
        PinWork,
        PinIdea,
        UnpinSelected
    ]
);

struct PasteOverlay {
    _platform_loop: Option<PlatformEventLoop>,
    focus_handle: FocusHandle,
    platform_status: String,
    platform: WindowsPlatformIntegration,
    history_path: Option<PathBuf>,
    document: ClipboardHistoryDocument,
    is_open: bool,
    is_trash_view: bool,
    is_plain_text: bool,
    is_command_palette_open: bool,
    command_palette_index: usize,
    is_search_active: bool,
    search_query: String,
    selected_index: usize,
    overlay_open_target: Option<WindowHandle>,
    overlay_window: Option<WindowHandle>,
    history_scroll: UniformListScrollHandle,
    animation_generation: u64,
    last_status: String,
    save_generation: u64,
}

impl PasteOverlay {
    fn new(
        platform_loop: Option<PlatformEventLoop>,
        platform_status: String,
        events: Option<Receiver<PlatformEvent>>,
        cx: &mut Context<Self>,
    ) -> Self {
        let history_path = default_history_path().ok();
        let (document, last_status) = match history_path.as_deref() {
            Some(path) => match load_history_document(path) {
                Ok(document) => (document, "History loaded".to_string()),
                Err(error) => (
                    create_document(seeded_groups(), []),
                    format!("History load failed: {error}"),
                ),
            },
            None => (
                create_document(seeded_groups(), []),
                "History path unavailable".to_string(),
            ),
        };

        if let Some(events) = events {
            cx.spawn(async move |this, cx| loop {
                let mut disconnected = false;
                loop {
                    match events.try_recv() {
                        Ok(event) => {
                            let _ = this.update(cx, |overlay, cx| {
                                overlay.handle_platform_event(event, cx);
                            });
                        }
                        Err(TryRecvError::Empty) => break,
                        Err(TryRecvError::Disconnected) => {
                            disconnected = true;
                            break;
                        }
                    }
                }

                if disconnected {
                    let _ = this.update(cx, |overlay, cx| {
                        overlay.last_status = "Platform event loop disconnected".to_string();
                        cx.notify();
                    });
                    break;
                }

                cx.background_executor()
                    .timer(Duration::from_millis(80))
                    .await;
            })
            .detach();
        }

        cx.observe_keystrokes(Self::observe_search_keystroke)
            .detach();
        cx.on_app_quit(|overlay, _cx| {
            overlay.flush_history();
            async {}
        })
        .detach();

        Self {
            _platform_loop: platform_loop,
            focus_handle: cx.focus_handle(),
            platform_status,
            platform: WindowsPlatformIntegration::new(),
            history_path,
            document,
            is_open: false,
            is_trash_view: false,
            is_plain_text: false,
            is_command_palette_open: false,
            command_palette_index: 0,
            is_search_active: false,
            search_query: String::new(),
            selected_index: 0,
            overlay_open_target: None,
            overlay_window: None,
            history_scroll: UniformListScrollHandle::new(),
            animation_generation: 0,
            last_status,
            save_generation: 0,
        }
    }

    fn handle_platform_event(&mut self, event: PlatformEvent, cx: &mut Context<Self>) {
        match event {
            PlatformEvent::OverlayHotkeyPressed => self.toggle_overlay(cx),
            PlatformEvent::ClipboardUpdated => self.capture_clipboard(cx),
            PlatformEvent::TrayShowRequested => {
                self.open_overlay("Tray requested overlay", cx);
            }
            PlatformEvent::TrayRestartRequested => {
                self.flush_history();
                match relaunch_current_exe() {
                    Ok(()) => {
                        self.last_status = "Restarting".to_string();
                        cx.quit();
                    }
                    Err(error) => {
                        self.last_status = format!("Restart failed: {error}");
                    }
                }
            }
            PlatformEvent::TrayExitRequested => {
                self.flush_history();
                self.last_status = "Tray exit requested".to_string();
                cx.quit();
            }
            PlatformEvent::MessageLoopStopped => {
                self.last_status = "Platform message loop stopped".to_string();
            }
        }
        cx.notify();
    }

    fn toggle_overlay(&mut self, cx: &mut Context<Self>) {
        self.is_open = !self.is_open;
        if self.is_open {
            self.overlay_open_target = self.platform.foreground_window();
            self.last_status = format!(
                "Overlay opened; target={}",
                self.overlay_open_target
                    .map(|window| format!("0x{:X}", window.0))
                    .unwrap_or_else(|| "none".to_string())
            );
            self.capture_clipboard(cx);
            self.activate_overlay_window(cx);
        } else {
            self.last_status = "Overlay hidden".to_string();
            self.animation_generation = self.animation_generation.wrapping_add(1);
            self.hide_native_overlay_window();
            cx.hide();
        }
    }

    fn open_overlay(&mut self, status: &str, cx: &mut Context<Self>) {
        self.is_open = true;
        self.overlay_open_target = self.platform.foreground_window();
        self.last_status = status.to_string();
        self.capture_clipboard(cx);
        self.activate_overlay_window(cx);
    }

    fn activate_overlay_window(&mut self, cx: &mut Context<Self>) {
        let focus_handle = self.focus_handle.clone();
        let entity_id = cx.entity_id();
        let native_window = self
            .overlay_window
            .or_else(|| self.find_native_overlay_window());
        if let Some(window) = native_window {
            self.overlay_window = Some(window);
            self.position_native_overlay_window(window, overlay_hidden_bounds(cx));
            self.show_native_overlay_window();
            self.animate_native_overlay_window(window, cx);
        }
        cx.with_window(entity_id, move |window, cx| {
            window.activate_window();
            window.focus(&focus_handle, cx);
        });
    }

    fn show_native_overlay_window(&self) {
        if let Some(window) = self
            .overlay_window
            .or_else(|| self.find_native_overlay_window())
        {
            let _ = self.platform.show_window(window);
        }
    }

    fn hide_native_overlay_window(&self) {
        if let Some(window) = self
            .overlay_window
            .or_else(|| self.find_native_overlay_window())
        {
            let _ = self.platform.hide_window(window);
        }
    }

    fn position_native_overlay_window(&self, window: WindowHandle, bounds: Bounds<Pixels>) {
        let _ = self.platform.move_window(
            window,
            bounds.origin.x.as_f32().round() as i32,
            bounds.origin.y.as_f32().round() as i32,
            bounds.size.width.as_f32().round() as i32,
            bounds.size.height.as_f32().round() as i32,
        );
    }

    fn animate_native_overlay_window(&mut self, window: WindowHandle, cx: &mut Context<Self>) {
        self.animation_generation = self.animation_generation.wrapping_add(1);
        let generation = self.animation_generation;

        for step in 0..=OVERLAY_ANIMATION_STEPS {
            cx.spawn(async move |this, cx| {
                cx.background_executor()
                    .timer(Duration::from_millis(OVERLAY_ANIMATION_FRAME_MS * step))
                    .await;
                let _ = this.update(cx, move |overlay, cx| {
                    if overlay.animation_generation != generation || !overlay.is_open {
                        return;
                    }
                    overlay
                        .position_native_overlay_window(window, animated_overlay_bounds(cx, step));
                });
            })
            .detach();
        }
    }

    fn find_native_overlay_window(&self) -> Option<WindowHandle> {
        self.platform
            .find_window_by_title(OVERLAY_WINDOW_TITLE)
            .ok()
            .flatten()
    }

    fn close_overlay(&mut self, _: &CloseOverlay, _window: &mut Window, cx: &mut Context<Self>) {
        if self.is_command_palette_open {
            self.is_command_palette_open = false;
            self.command_palette_index = 0;
            self.last_status = "Command palette closed".to_string();
            cx.notify();
            return;
        }

        self.is_open = false;
        self.last_status = "Overlay hidden".to_string();
        self.animation_generation = self.animation_generation.wrapping_add(1);
        self.hide_native_overlay_window();
        cx.hide();
        cx.notify();
    }

    fn move_up(&mut self, _: &MoveUp, _window: &mut Window, cx: &mut Context<Self>) {
        if self.is_command_palette_open {
            if self.command_palette_index > 0 {
                self.command_palette_index -= 1;
            }
            cx.notify();
            return;
        }

        if self.selected_index > 0 {
            self.selected_index -= 1;
        }
        cx.notify();
    }

    fn move_down(&mut self, _: &MoveDown, _window: &mut Window, cx: &mut Context<Self>) {
        if self.is_command_palette_open {
            let len = command_palette_commands(self.is_trash_view).len();
            if len > 0 {
                self.command_palette_index = (self.command_palette_index + 1).min(len - 1);
            }
            cx.notify();
            return;
        }

        let len = self.visible_entries().len();
        if len > 0 {
            self.selected_index = (self.selected_index + 1).min(len - 1);
        }
        cx.notify();
    }

    fn toggle_trash(&mut self, _: &ToggleTrash, _window: &mut Window, cx: &mut Context<Self>) {
        self.is_trash_view = !self.is_trash_view;
        self.selected_index = 0;
        self.last_status = if self.is_trash_view {
            "Trash view".to_string()
        } else {
            "History view".to_string()
        };
        cx.notify();
    }

    fn toggle_plain_text(
        &mut self,
        _: &TogglePlainText,
        _window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        self.is_plain_text = !self.is_plain_text;
        self.last_status = if self.is_plain_text {
            "Plain text paste enabled".to_string()
        } else {
            "Rich paste enabled".to_string()
        };
        cx.notify();
    }

    fn toggle_command_palette(
        &mut self,
        _: &ToggleCommandPalette,
        _window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        self.is_command_palette_open = !self.is_command_palette_open;
        self.command_palette_index = 0;
        self.last_status = if self.is_command_palette_open {
            "Command palette open".to_string()
        } else {
            "Command palette closed".to_string()
        };
        cx.notify();
    }

    fn toggle_search(&mut self, _: &ToggleSearch, _window: &mut Window, cx: &mut Context<Self>) {
        self.is_search_active = !self.is_search_active;
        if !self.is_search_active {
            self.search_query.clear();
        }
        self.selected_index = 0;
        self.last_status = if self.is_search_active {
            "Search active".to_string()
        } else {
            "Search inactive".to_string()
        };
        cx.notify();
    }

    fn observe_search_keystroke(
        &mut self,
        event: &gpui::KeystrokeEvent,
        _window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        if !self.is_search_active {
            return;
        }

        if event.action.is_some() {
            return;
        }

        match event.keystroke.key.as_str() {
            "backspace" => {
                self.search_query.pop();
                self.selected_index = 0;
                self.last_status = format!("Search: {}", self.search_query);
                cx.notify();
            }
            "escape" => {
                self.is_search_active = false;
                self.search_query.clear();
                self.selected_index = 0;
                self.last_status = "Search inactive".to_string();
                cx.notify();
            }
            _ => {
                if event.keystroke.modifiers.control
                    || event.keystroke.modifiers.alt
                    || event.keystroke.modifiers.platform
                    || event.keystroke.modifiers.function
                {
                    return;
                }

                let Some(text) = event.keystroke.key_char.as_deref() else {
                    return;
                };
                if text.chars().all(|ch| !ch.is_control()) {
                    self.search_query.push_str(text);
                    self.selected_index = 0;
                    self.last_status = format!("Search: {}", self.search_query);
                    cx.notify();
                }
            }
        }
    }

    fn paste_selected(&mut self, _: &PasteSelected, _window: &mut Window, cx: &mut Context<Self>) {
        if self.is_command_palette_open {
            self.execute_selected_command_palette_item(cx);
            cx.notify();
            return;
        }

        self.paste_selected_entry(false, cx);
        cx.notify();
    }

    fn one_shot_plain_paste(
        &mut self,
        _: &OneShotPlainPaste,
        _window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        self.paste_selected_entry(true, cx);
        cx.notify();
    }

    fn delete_selected(
        &mut self,
        _: &DeleteSelected,
        _window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        self.move_selected_to_trash(cx);
        cx.notify();
    }

    fn restore_selected(
        &mut self,
        _: &RestoreSelected,
        _window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        self.restore_selected_from_trash(cx);
        cx.notify();
    }

    fn pin_quick(&mut self, _: &PinQuick, _window: &mut Window, cx: &mut Context<Self>) {
        self.pin_selected_to_group(Some(QUICK_GROUP_ID), cx);
        cx.notify();
    }

    fn pin_work(&mut self, _: &PinWork, _window: &mut Window, cx: &mut Context<Self>) {
        self.pin_selected_to_group(Some(WORK_GROUP_ID), cx);
        cx.notify();
    }

    fn pin_idea(&mut self, _: &PinIdea, _window: &mut Window, cx: &mut Context<Self>) {
        self.pin_selected_to_group(Some(IDEA_GROUP_ID), cx);
        cx.notify();
    }

    fn unpin_selected(&mut self, _: &UnpinSelected, _window: &mut Window, cx: &mut Context<Self>) {
        self.pin_selected_to_group(None, cx);
        cx.notify();
    }

    fn pin_selected_to_group(&mut self, group_id: Option<&str>, cx: &mut Context<Self>) {
        let Some(selected) = self.selected_entry_identity() else {
            self.last_status = "No entry selected".to_string();
            return;
        };

        if let Some(entry) = self.find_entry_mut(&selected) {
            entry.pinned_group_id = group_id.map(ToString::to_string);
            self.last_status = match group_id {
                Some(group_id) => format!("Pinned selected entry to {group_id}"),
                None => "Unpinned selected entry".to_string(),
            };
            self.normalize_document_for_save();
            self.schedule_history_save(cx);
        }
    }

    fn execute_selected_command_palette_item(&mut self, cx: &mut Context<Self>) {
        let commands = command_palette_commands(self.is_trash_view);
        let Some(command) = commands.get(self.command_palette_index) else {
            self.last_status = "No command selected".to_string();
            return;
        };

        self.is_command_palette_open = false;
        self.command_palette_index = 0;
        match command.action {
            CommandPaletteAction::PasteSelected => self.paste_selected_entry(false, cx),
            CommandPaletteAction::PastePlainOnce => self.paste_selected_entry(true, cx),
            CommandPaletteAction::ToggleSearch => {
                self.is_search_active = !self.is_search_active;
                if !self.is_search_active {
                    self.search_query.clear();
                }
                self.selected_index = 0;
                self.last_status = if self.is_search_active {
                    "Search active".to_string()
                } else {
                    "Search inactive".to_string()
                };
            }
            CommandPaletteAction::ToggleTrash => {
                self.is_trash_view = !self.is_trash_view;
                self.selected_index = 0;
                self.last_status = if self.is_trash_view {
                    "Trash view".to_string()
                } else {
                    "History view".to_string()
                };
            }
            CommandPaletteAction::TogglePlainText => {
                self.is_plain_text = !self.is_plain_text;
                self.last_status = if self.is_plain_text {
                    "Plain text paste enabled".to_string()
                } else {
                    "Rich paste enabled".to_string()
                };
            }
            CommandPaletteAction::MoveToTrash => self.move_selected_to_trash(cx),
            CommandPaletteAction::RestoreFromTrash => self.restore_selected_from_trash(cx),
            CommandPaletteAction::PinQuick => self.pin_selected_to_group(Some(QUICK_GROUP_ID), cx),
            CommandPaletteAction::PinWork => self.pin_selected_to_group(Some(WORK_GROUP_ID), cx),
            CommandPaletteAction::PinIdea => self.pin_selected_to_group(Some(IDEA_GROUP_ID), cx),
            CommandPaletteAction::Unpin => self.pin_selected_to_group(None, cx),
            CommandPaletteAction::HideOverlay => {
                self.is_open = false;
                self.animation_generation = self.animation_generation.wrapping_add(1);
                self.last_status = "Overlay hidden".to_string();
                self.hide_native_overlay_window();
                cx.hide();
            }
        }
    }

    fn move_selected_to_trash(&mut self, cx: &mut Context<Self>) {
        if self.is_trash_view {
            self.last_status = "Already in trash; use Ctrl+R to restore".to_string();
            return;
        }

        let Some(selected) = self.selected_entry_identity() else {
            self.last_status = "No entry selected".to_string();
            return;
        };

        if let Some(entry) = self.find_entry_mut(&selected) {
            entry.is_deleted = true;
            entry.deleted_at = Some(OffsetDateTime::now_utc());
            self.last_status = "Moved selected entry to trash".to_string();
            self.normalize_document_for_save();
            self.clamp_selection();
            self.schedule_history_save(cx);
        }
    }

    fn restore_selected_from_trash(&mut self, cx: &mut Context<Self>) {
        if !self.is_trash_view {
            self.last_status = "Switch to trash to restore entries".to_string();
            return;
        }

        let Some(selected) = self.selected_entry_identity() else {
            self.last_status = "No trash entry selected".to_string();
            return;
        };

        if let Some(entry) = self.find_entry_mut(&selected) {
            entry.is_deleted = false;
            entry.deleted_at = None;
            self.last_status = "Restored selected entry".to_string();
            self.normalize_document_for_save();
            self.clamp_selection();
            self.schedule_history_save(cx);
        }
    }

    fn paste_selected_entry(&mut self, force_plain_text: bool, cx: &mut Context<Self>) {
        let visible = self.visible_entries();
        let Some(entry) = visible.get(self.selected_index) else {
            self.last_status = "No entry selected".to_string();
            return;
        };

        let paste_label = match entry.kind {
            paste_core::ClipboardKind::Text => {
                let text = entry.content.clone();
                if let Err(error) = self.platform.write_text_to_clipboard(&text) {
                    self.last_status = format!("Clipboard write failed: {error}");
                    return;
                }
                compact(&text, 40)
            }
            paste_core::ClipboardKind::Link => {
                let text = if self.is_plain_text || force_plain_text {
                    entry.content.clone()
                } else {
                    entry
                        .link_url
                        .clone()
                        .unwrap_or_else(|| entry.content.clone())
                };
                if let Err(error) = self.platform.write_text_to_clipboard(&text) {
                    self.last_status = format!("Clipboard write failed: {error}");
                    return;
                }
                compact(&text, 40)
            }
            paste_core::ClipboardKind::Image => {
                let Some(dib_bytes) = entry.image_dib_bytes.as_deref() else {
                    self.last_status = "Image paste unavailable; DIB bytes missing".to_string();
                    return;
                };
                if let Err(error) = self.platform.write_image_dib_to_clipboard(dib_bytes) {
                    self.last_status = format!("Clipboard image write failed: {error}");
                    return;
                }
                entry.content.clone()
            }
        };

        let target = self
            .overlay_open_target
            .or_else(|| self.platform.foreground_window());
        let Some(target) = target else {
            self.last_status = "Paste target unavailable".to_string();
            return;
        };

        self.is_open = false;
        self.animation_generation = self.animation_generation.wrapping_add(1);
        self.hide_native_overlay_window();
        cx.hide();
        match self.platform.paste_into_window(target) {
            Ok(()) => {
                self.last_status = format!("Pasted {paste_label}");
            }
            Err(error) => {
                self.last_status = format!("Paste injection failed: {error}");
            }
        }
    }

    fn capture_clipboard(&mut self, cx: &mut Context<Self>) {
        match self.platform.read_text_from_clipboard() {
            Ok(Some(text)) => {
                self.capture_clipboard_text(text, cx);
                return;
            }
            Ok(None) => {}
            Err(error) => {
                self.last_status = format!("Clipboard text read failed: {error}");
            }
        }

        match self.platform.read_image_dib_from_clipboard() {
            Ok(Some(bytes)) => self.capture_clipboard_image(bytes, cx),
            Ok(None) => {
                self.last_status = "Clipboard update ignored; no supported content".to_string();
            }
            Err(error) => {
                self.last_status = format!("Clipboard image read failed: {error}");
            }
        }
    }

    fn capture_clipboard_text(&mut self, text: String, cx: &mut Context<Self>) {
        if text.trim().is_empty() {
            self.last_status = "Clipboard update ignored; empty text".to_string();
            return;
        }

        let mut incoming = entry_from_clipboard_text(text, OffsetDateTime::now_utc());
        incoming.source_app = self.current_source_app_name();

        if let Some(existing) = self.document.entries.first_mut() {
            if is_likely_duplicate(existing, &incoming, Duration::from_secs(2)) {
                *existing = paste_core::merge_entries(existing, &incoming);
                self.last_status = "Merged duplicate clipboard text".to_string();
                self.normalize_document_for_save();
                self.schedule_history_save(cx);
                return;
            }
        }

        self.document.entries.insert(0, incoming);
        self.normalize_document_for_save();
        self.selected_index = 0;
        self.last_status = format!(
            "Captured text item; {} entries",
            self.document.entries.len()
        );
        self.schedule_history_save(cx);
    }

    fn capture_clipboard_image(&mut self, dib_bytes: Vec<u8>, cx: &mut Context<Self>) {
        let png_bytes = match dib_to_png_bytes(&dib_bytes) {
            Ok(bytes) => bytes,
            Err(error) => {
                self.last_status = format!("Clipboard image conversion failed: {error}");
                return;
            }
        };

        let mut incoming =
            entry_from_clipboard_image(png_bytes, dib_bytes, OffsetDateTime::now_utc());
        incoming.source_app = self.current_source_app_name();

        if let Some(existing) = self.document.entries.first_mut() {
            if is_likely_duplicate(existing, &incoming, Duration::from_secs(2)) {
                *existing = paste_core::merge_entries(existing, &incoming);
                self.last_status = "Merged duplicate clipboard image".to_string();
                self.normalize_document_for_save();
                self.schedule_history_save(cx);
                return;
            }
        }

        self.document.entries.insert(0, incoming);
        self.normalize_document_for_save();
        self.selected_index = 0;
        self.last_status = format!(
            "Captured image item; {} entries",
            self.document.entries.len()
        );
        self.schedule_history_save(cx);
    }

    fn schedule_history_save(&mut self, cx: &mut Context<Self>) {
        self.save_generation = self.save_generation.wrapping_add(1);
        let generation = self.save_generation;
        cx.spawn(async move |this, cx| {
            cx.background_executor().timer(HISTORY_SAVE_DEBOUNCE).await;
            let _ = this.update(cx, move |overlay, cx| {
                if overlay.save_generation == generation {
                    overlay.flush_history();
                    cx.notify();
                }
            });
        })
        .detach();
    }

    fn flush_history(&mut self) {
        self.save_generation = self.save_generation.wrapping_add(1);
        self.normalize_document_for_save();
        let Some(path) = self.history_path.as_deref() else {
            return;
        };

        if let Err(error) = save_history_document(path, &self.document) {
            self.last_status = format!("History save failed: {error}");
        }
    }

    fn normalize_document_for_save(&mut self) {
        let entries = enforce_retention(
            self.document.entries.clone(),
            OffsetDateTime::now_utc(),
            RetentionPolicy::default(),
        );
        self.document = create_document(self.document.groups.clone(), entries);
        self.clamp_selection();
    }

    fn current_source_app_name(&self) -> String {
        self.overlay_open_target
            .and_then(|window| self.platform.window_process_name(window).ok().flatten())
            .filter(|name| !name.trim().is_empty())
            .unwrap_or_else(|| "Clipboard".to_string())
    }

    fn visible_entries(&self) -> Vec<ClipboardEntry> {
        paste_core::visible_entries(
            &self.document.entries,
            self.is_trash_view,
            None,
            &self.search_query,
        )
    }

    fn selected_entry_identity(&self) -> Option<EntryIdentity> {
        self.visible_entries()
            .get(self.selected_index)
            .map(EntryIdentity::from)
    }

    fn find_entry_mut(&mut self, identity: &EntryIdentity) -> Option<&mut ClipboardEntry> {
        self.document.entries.iter_mut().find(|entry| {
            entry.copied_at == identity.copied_at
                && entry.kind == identity.kind
                && entry.content == identity.content
        })
    }

    fn clamp_selection(&mut self) {
        let len = self.visible_entries().len();
        if len == 0 {
            self.selected_index = 0;
        } else if self.selected_index >= len {
            self.selected_index = len - 1;
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct EntryIdentity {
    kind: paste_core::ClipboardKind,
    content: String,
    copied_at: OffsetDateTime,
}

impl From<&ClipboardEntry> for EntryIdentity {
    fn from(entry: &ClipboardEntry) -> Self {
        Self {
            kind: entry.kind.clone(),
            content: entry.content.clone(),
            copied_at: entry.copied_at,
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum CommandPaletteAction {
    PasteSelected,
    PastePlainOnce,
    ToggleSearch,
    ToggleTrash,
    TogglePlainText,
    MoveToTrash,
    RestoreFromTrash,
    PinQuick,
    PinWork,
    PinIdea,
    Unpin,
    HideOverlay,
}

#[derive(Clone, Copy, Debug)]
struct CommandPaletteCommand {
    label: &'static str,
    hint: &'static str,
    action: CommandPaletteAction,
}

fn command_palette_commands(is_trash_view: bool) -> Vec<CommandPaletteCommand> {
    let trash_command = if is_trash_view {
        CommandPaletteCommand {
            label: "Restore selected",
            hint: "Ctrl+R",
            action: CommandPaletteAction::RestoreFromTrash,
        }
    } else {
        CommandPaletteCommand {
            label: "Move selected to trash",
            hint: "Delete",
            action: CommandPaletteAction::MoveToTrash,
        }
    };

    vec![
        CommandPaletteCommand {
            label: "Paste selected",
            hint: "Enter",
            action: CommandPaletteAction::PasteSelected,
        },
        CommandPaletteCommand {
            label: "Paste as plain text once",
            hint: "Alt+Enter",
            action: CommandPaletteAction::PastePlainOnce,
        },
        CommandPaletteCommand {
            label: "Toggle search",
            hint: "Ctrl+F",
            action: CommandPaletteAction::ToggleSearch,
        },
        CommandPaletteCommand {
            label: if is_trash_view {
                "Show history"
            } else {
                "Show trash"
            },
            hint: "Ctrl+T",
            action: CommandPaletteAction::ToggleTrash,
        },
        CommandPaletteCommand {
            label: "Toggle plain text",
            hint: "Ctrl+Shift+V",
            action: CommandPaletteAction::TogglePlainText,
        },
        trash_command,
        CommandPaletteCommand {
            label: "Pin to Quick",
            hint: "Ctrl+1",
            action: CommandPaletteAction::PinQuick,
        },
        CommandPaletteCommand {
            label: "Pin to Work",
            hint: "Ctrl+2",
            action: CommandPaletteAction::PinWork,
        },
        CommandPaletteCommand {
            label: "Pin to Idea",
            hint: "Ctrl+3",
            action: CommandPaletteAction::PinIdea,
        },
        CommandPaletteCommand {
            label: "Unpin selected",
            hint: "Ctrl+0",
            action: CommandPaletteAction::Unpin,
        },
        CommandPaletteCommand {
            label: "Hide overlay",
            hint: "Esc",
            action: CommandPaletteAction::HideOverlay,
        },
    ]
}

fn render_entry_row(entry: &ClipboardEntry, selected: bool) -> impl IntoElement {
    div()
        .id(format!(
            "entry-{}-{}",
            entry.kind_label(),
            entry.copied_at.unix_timestamp_nanos()
        ))
        .flex()
        .flex_col()
        .gap_1()
        .mb(px(8.0))
        .p(px(10.0))
        .rounded_sm()
        .text_size(px(13.0))
        .bg(if selected {
            rgb(0x2D4F67)
        } else {
            rgb(0x20242B)
        })
        .child(format!(
            "{}{}  {}",
            entry.kind_label(),
            entry
                .pinned_group_id
                .as_deref()
                .map(|group| format!(" [{}]", group))
                .unwrap_or_default(),
            compact(&entry.content, 72)
        ))
        .child(format!(
            "{}  {}",
            compact(&entry.source_app, 24),
            copied_at_label(entry.copied_at)
        ))
}

trait EntryKindLabel {
    fn kind_label(&self) -> &'static str;
}

impl EntryKindLabel for ClipboardEntry {
    fn kind_label(&self) -> &'static str {
        match self.kind {
            paste_core::ClipboardKind::Text => "Text",
            paste_core::ClipboardKind::Image => "Image",
            paste_core::ClipboardKind::Link => "Link",
        }
    }
}

impl Focusable for PasteOverlay {
    fn focus_handle(&self, _cx: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl Render for PasteOverlay {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        window.set_window_title(OVERLAY_WINDOW_TITLE);
        if self.overlay_window.is_none() {
            self.overlay_window = self.find_native_overlay_window();
        }
        if !self.is_open {
            self.hide_native_overlay_window();
        }

        let visible_entries = self.visible_entries();
        if !visible_entries.is_empty() {
            self.history_scroll.scroll_to_item(
                self.selected_index.min(visible_entries.len() - 1),
                ScrollStrategy::Nearest,
            );
        }
        let selected_index = self.selected_index;
        let preview_entries = visible_entries.clone();
        let preview = uniform_list(
            "paste-history-list",
            preview_entries.len(),
            move |range, _window, _cx| {
                range
                    .filter_map(|index| preview_entries.get(index).map(|entry| (index, entry)))
                    .map(|(index, entry)| render_entry_row(entry, index == selected_index))
                    .collect::<Vec<_>>()
            },
        )
        .track_scroll(&self.history_scroll)
        .h(px(430.0));

        let header = div()
            .flex()
            .justify_between()
            .text_size(px(14.0))
            .child(format!(
                "{} items  {}  {}",
                visible_entries.len(),
                if self.is_trash_view {
                    "Trash"
                } else {
                    "History"
                },
                if self.is_plain_text { "Plain" } else { "Rich" }
            ))
            .child(if self.is_search_active {
                format!("Search: {}", self.search_query)
            } else {
                "Search idle".to_string()
            });

        let command_palette_commands = command_palette_commands(self.is_trash_view);
        let command_palette = command_palette_commands.iter().enumerate().fold(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .p(px(10.0))
                .rounded_sm()
                .text_size(px(13.0))
                .bg(rgb(0x1B2028))
                .child("Command palette"),
            |palette, (index, command)| {
                let selected = self.command_palette_index == index;
                palette.child(
                    div()
                        .flex()
                        .justify_between()
                        .px(px(8.0))
                        .py(px(5.0))
                        .rounded_sm()
                        .bg(if selected {
                            rgb(0x2D4F67)
                        } else {
                            rgb(0x1B2028)
                        })
                        .child(command.label)
                        .child(command.hint),
                )
            },
        );

        div()
            .id("paste-overlay")
            .key_context("PasteOverlay")
            .track_focus(&self.focus_handle)
            .on_action(cx.listener(Self::move_up))
            .on_action(cx.listener(Self::move_down))
            .on_action(cx.listener(Self::paste_selected))
            .on_action(cx.listener(Self::one_shot_plain_paste))
            .on_action(cx.listener(Self::close_overlay))
            .on_action(cx.listener(Self::toggle_trash))
            .on_action(cx.listener(Self::toggle_plain_text))
            .on_action(cx.listener(Self::toggle_command_palette))
            .on_action(cx.listener(Self::toggle_search))
            .on_action(cx.listener(Self::delete_selected))
            .on_action(cx.listener(Self::restore_selected))
            .on_action(cx.listener(Self::pin_quick))
            .on_action(cx.listener(Self::pin_work))
            .on_action(cx.listener(Self::pin_idea))
            .on_action(cx.listener(Self::unpin_selected))
            .flex()
            .flex_col()
            .gap_3()
            .size_full()
            .rounded_lg()
            .border_1()
            .border_color(rgb(0x303844))
            .bg(rgb(0x111318))
            .text_color(rgb(0xE8EAED))
            .text_size(px(13.0))
            .p(px(24.0))
            .child("Paste GPUI")
            .child(header)
            .child(self.platform_status.clone())
            .child(self.last_status.clone())
            .when(self.is_command_palette_open, |view| {
                view.child(command_palette)
            })
            .child(preview)
    }
}

fn compact(value: &str, max_chars: usize) -> String {
    let mut compacted = value.replace("\r\n", " ").replace('\n', " ");
    if compacted.chars().count() > max_chars {
        compacted = compacted
            .chars()
            .take(max_chars.saturating_sub(1))
            .collect();
        compacted.push_str("...");
    }
    compacted
}

fn copied_at_label(copied_at: OffsetDateTime) -> String {
    copied_at
        .format(&time::macros::format_description!(
            "[year]-[month]-[day] [hour]:[minute]:[second]"
        ))
        .unwrap_or_else(|_| "unknown time".to_string())
}

fn relaunch_current_exe() -> io::Result<()> {
    let exe = env::current_exe()?;
    let mut command = Command::new(exe);
    command.args(env::args_os().skip(1));
    command.spawn().map(|_| ())
}

fn entry_from_clipboard_text(text: String, copied_at: OffsetDateTime) -> ClipboardEntry {
    let trimmed = text.trim();
    if let Ok(url) = Url::parse(trimmed) {
        if matches!(url.scheme(), "http" | "https") {
            let mut entry = ClipboardEntry::text(text.clone(), copied_at);
            entry.kind = paste_core::ClipboardKind::Link;
            entry.content = text;
            entry.link_url = Some(url.as_str().to_string());
            entry.link_title = None;
            return entry;
        }
    }

    ClipboardEntry::text(text, copied_at)
}

fn entry_from_clipboard_image(
    png_bytes: Vec<u8>,
    dib_bytes: Vec<u8>,
    copied_at: OffsetDateTime,
) -> ClipboardEntry {
    let mut entry = ClipboardEntry::text(format!("[Image: {} bytes]", png_bytes.len()), copied_at);
    entry.kind = paste_core::ClipboardKind::Image;
    entry.image_png_bytes = Some(png_bytes);
    entry.image_dib_bytes = Some(dib_bytes);
    entry
}

fn start_platform() -> (
    Option<PlatformEventLoop>,
    Option<Receiver<PlatformEvent>>,
    String,
) {
    match spawn_platform_event_loop() {
        Ok((platform_loop, events)) => {
            let window = platform_loop.window();
            (
                Some(platform_loop),
                Some(events),
                format!("Win32 event loop active on hidden window 0x{:X}", window.0),
            )
        }
        Err(error) => (None, None, format!("Win32 event loop failed: {error}")),
    }
}

fn overlay_window_options(cx: &App) -> WindowOptions {
    WindowOptions {
        titlebar: None,
        focus: false,
        show: true,
        is_resizable: false,
        is_minimizable: false,
        kind: WindowKind::Floating,
        window_bounds: Some(WindowBounds::Windowed(bottom_overlay_bounds(cx))),
        ..Default::default()
    }
}

fn bottom_overlay_bounds(cx: &App) -> Bounds<Pixels> {
    let panel_size = overlay_panel_size();
    let Some(display) = cx.primary_display() else {
        return Bounds::centered(None, panel_size, cx);
    };

    let display_bounds = display.bounds();
    let x = display_bounds.origin.x + (display_bounds.size.width - panel_size.width) / 2.0;
    let y = display_bounds.origin.y + display_bounds.size.height
        - panel_size.height
        - px(OVERLAY_BOTTOM_MARGIN);
    bounds(point(x, y), panel_size)
}

fn overlay_hidden_bounds(cx: &App) -> Bounds<Pixels> {
    let visible = bottom_overlay_bounds(cx);
    bounds(
        point(
            visible.origin.x,
            visible.origin.y + visible.size.height + px(16.0),
        ),
        visible.size,
    )
}

fn animated_overlay_bounds(cx: &App, step: u64) -> Bounds<Pixels> {
    let visible = bottom_overlay_bounds(cx);
    let hidden = overlay_hidden_bounds(cx);
    let progress = (step.min(OVERLAY_ANIMATION_STEPS) as f32) / (OVERLAY_ANIMATION_STEPS as f32);
    let eased = 1.0 - (1.0 - progress).powi(3);
    let y =
        hidden.origin.y.as_f32() + (visible.origin.y.as_f32() - hidden.origin.y.as_f32()) * eased;
    bounds(point(visible.origin.x, px(y)), visible.size)
}

fn overlay_panel_size() -> gpui::Size<Pixels> {
    size(px(OVERLAY_PANEL_WIDTH), px(OVERLAY_PANEL_HEIGHT))
}

fn main() {
    let (platform_loop, events, platform_status) = start_platform();

    application().run(|cx: &mut App| {
        cx.bind_keys([
            KeyBinding::new("up", MoveUp, Some("PasteOverlay")),
            KeyBinding::new("down", MoveDown, Some("PasteOverlay")),
            KeyBinding::new("enter", PasteSelected, Some("PasteOverlay")),
            KeyBinding::new("alt-enter", OneShotPlainPaste, Some("PasteOverlay")),
            KeyBinding::new("escape", CloseOverlay, Some("PasteOverlay")),
            KeyBinding::new("ctrl-t", ToggleTrash, Some("PasteOverlay")),
            KeyBinding::new("ctrl-k", ToggleCommandPalette, Some("PasteOverlay")),
            KeyBinding::new("ctrl-f", ToggleSearch, Some("PasteOverlay")),
            KeyBinding::new("ctrl-shift-v", TogglePlainText, Some("PasteOverlay")),
            KeyBinding::new("delete", DeleteSelected, Some("PasteOverlay")),
            KeyBinding::new("ctrl-r", RestoreSelected, Some("PasteOverlay")),
            KeyBinding::new("ctrl-1", PinQuick, Some("PasteOverlay")),
            KeyBinding::new("ctrl-2", PinWork, Some("PasteOverlay")),
            KeyBinding::new("ctrl-3", PinIdea, Some("PasteOverlay")),
            KeyBinding::new("ctrl-0", UnpinSelected, Some("PasteOverlay")),
        ]);
        cx.open_window(overlay_window_options(cx), |_window, cx| {
            cx.new(|cx| PasteOverlay::new(platform_loop, platform_status, events, cx))
        })
        .expect("open GPUI feasibility gate window");
    });
}

#[cfg(test)]
mod tests {
    use super::*;

    fn at(seconds: i64) -> OffsetDateTime {
        OffsetDateTime::from_unix_timestamp(seconds).unwrap()
    }

    #[test]
    fn clipboard_url_becomes_link_entry() {
        let entry = entry_from_clipboard_text("https://example.com/path".to_string(), at(0));

        assert_eq!(entry.kind, paste_core::ClipboardKind::Link);
        assert_eq!(entry.link_url.as_deref(), Some("https://example.com/path"));
    }

    #[test]
    fn clipboard_non_url_stays_text_entry() {
        let entry = entry_from_clipboard_text("not a url".to_string(), at(0));

        assert_eq!(entry.kind, paste_core::ClipboardKind::Text);
        assert_eq!(entry.link_url, None);
    }

    #[test]
    fn clipboard_image_becomes_image_entry() {
        let entry = entry_from_clipboard_image(vec![1, 2, 3], vec![4, 5, 6], at(0));

        assert_eq!(entry.kind, paste_core::ClipboardKind::Image);
        assert_eq!(entry.content, "[Image: 3 bytes]");
        assert_eq!(entry.image_png_bytes.as_deref(), Some(&[1, 2, 3][..]));
        assert_eq!(entry.image_dib_bytes.as_deref(), Some(&[4, 5, 6][..]));
    }
}
