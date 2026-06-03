use gpui::prelude::*;
use gpui::{
    actions, bounds, div, hsla, linear_color_stop, linear_gradient, point, px, rgb, rgba, size,
    App, Bounds, BoxShadow, Context, FocusHandle, Focusable, IntoElement, KeyBinding, Pixels,
    Render, Window, WindowBackgroundAppearance, WindowBounds, WindowDecorations, WindowKind,
    WindowOptions,
};
use gpui_platform::application;
use paste_core::{
    create_document, default_history_path, dib_to_png_bytes, enforce_retention,
    is_likely_duplicate, load_history_document, save_history_document, seeded_groups,
    visible_favorite_entries, visible_kind_entries, ClipboardEntry, ClipboardHistoryDocument,
    ClipboardKind, RetentionPolicy, IDEA_GROUP_ID, QUICK_GROUP_ID, WORK_GROUP_ID,
};
use paste_windows_platform::{
    spawn_platform_event_loop, PlatformEvent, PlatformEventLoop, PlatformIntegration, WindowHandle,
    WindowsPlatformIntegration,
};
use std::{
    env, io,
    path::{Path, PathBuf},
    process::Command,
    sync::mpsc::{Receiver, TryRecvError},
    time::Duration,
};
use time::OffsetDateTime;
use url::Url;

const HISTORY_SAVE_DEBOUNCE: Duration = Duration::from_millis(300);
const OVERLAY_WINDOW_TITLE: &str = "Paste GPUI Overlay";
const OVERLAY_FALLBACK_WIDTH: f32 = 1120.0;
const OVERLAY_SCREEN_WIDTH_RATIO: f32 = 0.90;
const OVERLAY_SAFE_SIDE_MARGIN: f32 = 32.0;
const OVERLAY_WINDOW_HEIGHT: f32 = 356.0;
const MAIN_PANEL_HEIGHT: f32 = 356.0;
const PREVIEW_POPOVER_HEIGHT: f32 = 88.0;
const OVERLAY_BOTTOM_MARGIN: f32 = 32.0;
const OVERLAY_ANIMATION_STEPS: u64 = 8;
const OVERLAY_ANIMATION_FRAME_MS: u64 = 14;
const SELECTION_ANIMATION_STEPS: u64 = 6;
const SELECTION_ANIMATION_FRAME_MS: u64 = 18;
const PANEL_RADIUS: f32 = 28.0;
const PANEL_BACKGROUND_ALPHA: u32 = 0xE6;
const PANEL_GLASS_HIGHLIGHT_ALPHA: u32 = 0x24;
const PANEL_GLASS_EDGE_ALPHA: u32 = 0x18;
const PANEL_GLASS_SHADE_ALPHA: u32 = 0x22;
const CARD_RADIUS: f32 = 20.0;
const PANEL_HORIZONTAL_PADDING: f32 = 28.0;
const RAIL_CARD_WIDTH: f32 = 180.0;
const RAIL_CARD_HEIGHT: f32 = 220.0;
const RAIL_CARD_GAP: f32 = 14.0;
const THEME_PURPLE: u32 = 0x7C5CFF;
const THEME_SELECTION_BLUE: u32 = 0x0A84FF;
const THEME_SUCCESS: u32 = 0x3DDC97;
const THEME_WARNING: u32 = 0xFFD166;
const THEME_ERROR: u32 = 0xFF5A5F;
const THEME_BLUE: u32 = 0x4DA3FF;
const THEME_ORANGE: u32 = 0xFF8A5B;
const THEME_PINBOARD: u32 = THEME_WARNING;

#[derive(Clone, Copy)]
struct AppTheme {
    panel: u32,
    panel_alt: u32,
    surface: u32,
    card: u32,
    card_hover: u32,
    card_text: u32,
    text: u32,
    muted: u32,
    weak: u32,
    border: u32,
    key_bg: u32,
}

actions!(
    paste_overlay,
    [
        MoveUp,
        MoveDown,
        MoveLeft,
        MoveRight,
        PasteSelected,
        CloseOverlay,
        ToggleTrash,
        TogglePlainText,
        ToggleCommandPalette,
        ToggleSearch,
        ToggleFavorite,
        ToggleContextMenu,
        ToggleSettings,
        ConfirmDelete,
        CancelDelete,
        OneShotPlainPaste,
        DeleteSelected,
        RestoreSelected,
        PinQuick,
        PinWork,
        PinIdea,
        UnpinSelected
    ]
);

#[derive(Clone, Debug, Eq, PartialEq)]
enum OverlayFilter {
    All,
    Favorites,
    Kind(ClipboardKind),
    Group(String),
}

impl OverlayFilter {
    fn label(&self) -> String {
        match self {
            Self::All => "All".to_string(),
            Self::Favorites => "Favorites".to_string(),
            Self::Kind(kind) => kind_label(kind).to_string(),
            Self::Group(group_id) => group_id.to_string(),
        }
    }
}

struct PasteOverlay {
    _platform_loop: Option<PlatformEventLoop>,
    focus_handle: FocusHandle,
    platform: WindowsPlatformIntegration,
    history_path: Option<PathBuf>,
    document: ClipboardHistoryDocument,
    is_open: bool,
    is_trash_view: bool,
    is_plain_text: bool,
    is_command_palette_open: bool,
    command_palette_index: usize,
    is_context_menu_open: bool,
    is_settings_open: bool,
    is_search_active: bool,
    search_query: String,
    active_filter: OverlayFilter,
    selected_index: usize,
    pending_delete: Option<EntryIdentity>,
    overlay_open_target: Option<WindowHandle>,
    overlay_window: Option<WindowHandle>,
    animation_generation: u64,
    selection_animation_generation: u64,
    selection_pulse: f32,
    reduced_motion: bool,
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

        let platform = WindowsPlatformIntegration::new();
        let reduced_motion = reduced_motion_requested(&platform);
        let last_status = format!("{last_status}; {platform_status}");

        Self {
            _platform_loop: platform_loop,
            focus_handle: cx.focus_handle(),
            platform,
            history_path,
            document,
            is_open: false,
            is_trash_view: false,
            is_plain_text: false,
            is_command_palette_open: false,
            command_palette_index: 0,
            is_context_menu_open: env_flag_enabled("PASTE_GPUI_SHOW_CONTEXT_MENU"),
            is_settings_open: env_flag_enabled("PASTE_GPUI_SHOW_SETTINGS"),
            is_search_active: false,
            search_query: String::new(),
            active_filter: OverlayFilter::All,
            selected_index: 0,
            pending_delete: None,
            overlay_open_target: None,
            overlay_window: None,
            animation_generation: 0,
            selection_animation_generation: 0,
            selection_pulse: 0.0,
            reduced_motion,
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
            let _ = self.platform.suppress_window_border(window);
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
            let _ = self.platform.suppress_window_border(window);
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
        let width = bounds.size.width.as_f32().round() as i32;
        let height = bounds.size.height.as_f32().round() as i32;
        let _ = self.platform.move_window(
            window,
            bounds.origin.x.as_f32().round() as i32,
            bounds.origin.y.as_f32().round() as i32,
            width,
            height,
        );
        let _ = self.platform.apply_overlay_window_shape(
            window,
            width,
            height,
            PANEL_RADIUS.round() as i32,
        );
        let _ = self.platform.suppress_window_border(window);
    }

    fn animate_native_overlay_window(&mut self, window: WindowHandle, cx: &mut Context<Self>) {
        self.animation_generation = self.animation_generation.wrapping_add(1);
        let generation = self.animation_generation;

        if self.reduced_motion {
            self.position_native_overlay_window(window, bottom_overlay_bounds(cx));
            return;
        }

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
        if self.pending_delete.take().is_some() {
            self.last_status = "Delete cancelled".to_string();
            cx.notify();
            return;
        }

        if self.is_command_palette_open {
            self.is_command_palette_open = false;
            self.command_palette_index = 0;
            self.last_status = "Command palette closed".to_string();
            cx.notify();
            return;
        }

        if self.is_context_menu_open {
            self.is_context_menu_open = false;
            self.last_status = "Context menu closed".to_string();
            cx.notify();
            return;
        }

        if self.is_settings_open {
            self.is_settings_open = false;
            self.last_status = "Settings closed".to_string();
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

        if self.move_selection_left() {
            self.start_selection_animation(cx);
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

        if self.move_selection_right() {
            self.start_selection_animation(cx);
        }
        cx.notify();
    }

    fn move_left(&mut self, _: &MoveLeft, _window: &mut Window, cx: &mut Context<Self>) {
        if self.is_command_palette_open {
            if self.command_palette_index > 0 {
                self.command_palette_index -= 1;
            }
            cx.notify();
            return;
        }

        if self.move_selection_left() {
            self.start_selection_animation(cx);
        }
        cx.notify();
    }

    fn move_right(&mut self, _: &MoveRight, _window: &mut Window, cx: &mut Context<Self>) {
        if self.is_command_palette_open {
            let len = command_palette_commands(self.is_trash_view).len();
            if len > 0 {
                self.command_palette_index = (self.command_palette_index + 1).min(len - 1);
            }
            cx.notify();
            return;
        }

        if self.move_selection_right() {
            self.start_selection_animation(cx);
        }
        cx.notify();
    }

    fn move_selection_left(&mut self) -> bool {
        if self.selected_index > 0 {
            self.selected_index -= 1;
            self.pending_delete = None;
            true
        } else {
            false
        }
    }

    fn move_selection_right(&mut self) -> bool {
        let len = self.visible_entries().len();
        if len > 0 {
            let next = (self.selected_index + 1).min(len - 1);
            if next != self.selected_index {
                self.selected_index = next;
                self.pending_delete = None;
                return true;
            }
        }
        false
    }

    fn toggle_trash(&mut self, _: &ToggleTrash, _window: &mut Window, cx: &mut Context<Self>) {
        self.is_trash_view = !self.is_trash_view;
        self.selected_index = 0;
        self.pending_delete = None;
        self.last_status = if self.is_trash_view {
            "Trash view".to_string()
        } else {
            "History view".to_string()
        };
        self.start_selection_animation(cx);
        cx.notify();
    }

    fn toggle_plain_text(
        &mut self,
        _: &TogglePlainText,
        _window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        self.is_plain_text = !self.is_plain_text;
        self.pending_delete = None;
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
        self.is_context_menu_open = false;
        self.pending_delete = None;
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
        self.pending_delete = None;
        self.last_status = if self.is_search_active {
            "Search active".to_string()
        } else {
            "Search inactive".to_string()
        };
        self.start_selection_animation(cx);
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
                self.pending_delete = None;
                self.last_status = format!("Search: {}", self.search_query);
                self.start_selection_animation(cx);
                cx.notify();
            }
            "escape" => {
                self.is_search_active = false;
                self.search_query.clear();
                self.selected_index = 0;
                self.pending_delete = None;
                self.last_status = "Search inactive".to_string();
                self.start_selection_animation(cx);
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
                    self.pending_delete = None;
                    self.last_status = format!("Search: {}", self.search_query);
                    self.start_selection_animation(cx);
                    cx.notify();
                }
            }
        }
    }

    fn paste_selected(&mut self, _: &PasteSelected, _window: &mut Window, cx: &mut Context<Self>) {
        if self.pending_delete.is_some() {
            self.confirm_pending_delete(cx);
            cx.notify();
            return;
        }

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
        if self.pending_delete.is_some() {
            self.confirm_pending_delete(cx);
            cx.notify();
            return;
        }

        self.paste_selected_entry(true, cx);
        cx.notify();
    }

    fn toggle_favorite(
        &mut self,
        _: &ToggleFavorite,
        _window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        self.toggle_selected_favorite(cx);
        cx.notify();
    }

    fn toggle_context_menu(
        &mut self,
        _: &ToggleContextMenu,
        _window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        self.is_context_menu_open = !self.is_context_menu_open;
        self.is_command_palette_open = false;
        self.is_settings_open = false;
        self.pending_delete = None;
        self.last_status = if self.is_context_menu_open {
            "Context menu open".to_string()
        } else {
            "Context menu closed".to_string()
        };
        cx.notify();
    }

    fn toggle_settings(
        &mut self,
        _: &ToggleSettings,
        _window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        self.is_settings_open = !self.is_settings_open;
        self.is_command_palette_open = false;
        self.is_context_menu_open = false;
        self.pending_delete = None;
        self.last_status = if self.is_settings_open {
            "Settings open".to_string()
        } else {
            "Settings closed".to_string()
        };
        cx.notify();
    }

    fn confirm_delete(&mut self, _: &ConfirmDelete, _window: &mut Window, cx: &mut Context<Self>) {
        self.confirm_pending_delete(cx);
        cx.notify();
    }

    fn cancel_delete(&mut self, _: &CancelDelete, _window: &mut Window, cx: &mut Context<Self>) {
        if self.pending_delete.take().is_some() {
            self.last_status = "Delete cancelled".to_string();
        }
        cx.notify();
    }

    fn delete_selected(
        &mut self,
        _: &DeleteSelected,
        _window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        self.request_delete_selected();
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
            self.pending_delete = None;
            self.last_status = match group_id {
                Some(group_id) => format!("Pinned selected entry to {group_id}"),
                None => "Unpinned selected entry".to_string(),
            };
            self.normalize_document_for_save();
            self.schedule_history_save(cx);
        }
    }

    fn toggle_selected_favorite(&mut self, cx: &mut Context<Self>) {
        let Some(selected) = self.selected_entry_identity() else {
            self.last_status = "No entry selected".to_string();
            return;
        };

        if let Some(entry) = self.find_entry_mut(&selected) {
            entry.is_favorite = !entry.is_favorite;
            let is_favorite = entry.is_favorite;
            self.pending_delete = None;
            self.last_status = if is_favorite {
                "Added selected entry to Favorites".to_string()
            } else {
                "Removed selected entry from Favorites".to_string()
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
        let action = command.action;
        let animate_selection = command_action_changes_selection_surface(action);
        match action {
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
            CommandPaletteAction::MoveToTrash => self.request_delete_selected(),
            CommandPaletteAction::RestoreFromTrash => self.restore_selected_from_trash(cx),
            CommandPaletteAction::PinQuick => self.pin_selected_to_group(Some(QUICK_GROUP_ID), cx),
            CommandPaletteAction::PinWork => self.pin_selected_to_group(Some(WORK_GROUP_ID), cx),
            CommandPaletteAction::PinIdea => self.pin_selected_to_group(Some(IDEA_GROUP_ID), cx),
            CommandPaletteAction::Unpin => self.pin_selected_to_group(None, cx),
            CommandPaletteAction::ToggleFavorite => self.toggle_selected_favorite(cx),
            CommandPaletteAction::ShowAll => self.set_filter(OverlayFilter::All),
            CommandPaletteAction::ShowFavorites => self.set_filter(OverlayFilter::Favorites),
            CommandPaletteAction::FilterText => {
                self.set_filter(OverlayFilter::Kind(ClipboardKind::Text))
            }
            CommandPaletteAction::FilterLink => {
                self.set_filter(OverlayFilter::Kind(ClipboardKind::Link))
            }
            CommandPaletteAction::FilterImage => {
                self.set_filter(OverlayFilter::Kind(ClipboardKind::Image))
            }
            CommandPaletteAction::FilterFile => {
                self.set_filter(OverlayFilter::Kind(ClipboardKind::File))
            }
            CommandPaletteAction::FilterCode => {
                self.set_filter(OverlayFilter::Kind(ClipboardKind::Code))
            }
            CommandPaletteAction::ShowQuick => {
                self.set_filter(OverlayFilter::Group(QUICK_GROUP_ID.to_string()))
            }
            CommandPaletteAction::ShowWork => {
                self.set_filter(OverlayFilter::Group(WORK_GROUP_ID.to_string()))
            }
            CommandPaletteAction::ShowIdea => {
                self.set_filter(OverlayFilter::Group(IDEA_GROUP_ID.to_string()))
            }
            CommandPaletteAction::ToggleContextMenu => {
                self.is_context_menu_open = !self.is_context_menu_open;
                self.is_settings_open = false;
                self.last_status = if self.is_context_menu_open {
                    "Context menu open".to_string()
                } else {
                    "Context menu closed".to_string()
                };
            }
            CommandPaletteAction::ToggleSettings => {
                self.is_settings_open = !self.is_settings_open;
                self.is_context_menu_open = false;
                self.last_status = if self.is_settings_open {
                    "Settings open".to_string()
                } else {
                    "Settings closed".to_string()
                };
            }
            CommandPaletteAction::HideOverlay => {
                self.is_open = false;
                self.animation_generation = self.animation_generation.wrapping_add(1);
                self.last_status = "Overlay hidden".to_string();
                self.hide_native_overlay_window();
                cx.hide();
            }
        }

        if animate_selection && self.is_open {
            self.start_selection_animation(cx);
        }
    }

    fn set_filter(&mut self, filter: OverlayFilter) {
        self.active_filter = filter;
        self.selected_index = 0;
        self.pending_delete = None;
        self.last_status = format!("Filter: {}", self.active_filter.label());
        self.clamp_selection();
    }

    fn start_selection_animation(&mut self, cx: &mut Context<Self>) {
        self.selection_animation_generation = self.selection_animation_generation.wrapping_add(1);
        let generation = self.selection_animation_generation;

        if self.reduced_motion {
            self.selection_pulse = 0.0;
            return;
        }

        self.selection_pulse = 1.0;
        for step in 1..=SELECTION_ANIMATION_STEPS {
            cx.spawn(async move |this, cx| {
                cx.background_executor()
                    .timer(Duration::from_millis(SELECTION_ANIMATION_FRAME_MS * step))
                    .await;
                let _ = this.update(cx, move |overlay, cx| {
                    if overlay.selection_animation_generation != generation {
                        return;
                    }
                    overlay.selection_pulse = selection_pulse_for_step(step);
                    cx.notify();
                });
            })
            .detach();
        }
    }

    fn request_delete_selected(&mut self) {
        if self.is_trash_view {
            self.last_status = "Already in trash; use Ctrl+R to restore".to_string();
            return;
        }

        let Some(selected) = self.selected_entry_identity() else {
            self.last_status = "No entry selected".to_string();
            return;
        };

        self.pending_delete = Some(selected);
        self.is_context_menu_open = false;
        self.is_settings_open = false;
        self.last_status = "Confirm delete with Enter; Esc cancels".to_string();
    }

    fn confirm_pending_delete(&mut self, cx: &mut Context<Self>) {
        let Some(identity) = self.pending_delete.take() else {
            self.last_status = "No delete pending".to_string();
            return;
        };

        self.move_identity_to_trash(identity, cx);
    }

    fn move_identity_to_trash(&mut self, identity: EntryIdentity, cx: &mut Context<Self>) {
        if let Some(entry) = self.find_entry_mut(&identity) {
            entry.is_deleted = true;
            entry.deleted_at = Some(OffsetDateTime::now_utc());
            self.last_status = "Moved selected entry to trash".to_string();
            self.normalize_document_for_save();
            self.clamp_selection();
            self.schedule_history_save(cx);
        } else {
            self.last_status = "Delete target no longer exists".to_string();
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
            self.pending_delete = None;
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

        let paste_label = match &entry.kind {
            ClipboardKind::Text => {
                let text = entry.content.clone();
                if let Err(error) = self.platform.write_text_to_clipboard(&text) {
                    self.last_status = format!("Clipboard write failed: {error}");
                    return;
                }
                compact(&text, 40)
            }
            ClipboardKind::Link => {
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
            ClipboardKind::Image => {
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
            ClipboardKind::File => {
                if entry.file_paths.is_empty() {
                    self.last_status = "File paste unavailable; file paths missing".to_string();
                    return;
                }
                if self.is_plain_text || force_plain_text {
                    let text = entry.file_paths.join("\r\n");
                    if let Err(error) = self.platform.write_text_to_clipboard(&text) {
                        self.last_status = format!("Clipboard write failed: {error}");
                        return;
                    }
                } else if let Err(error) = self
                    .platform
                    .write_file_paths_to_clipboard(&entry.file_paths)
                {
                    self.last_status = format!("Clipboard file write failed: {error}");
                    return;
                }
                file_entry_label(entry)
            }
            ClipboardKind::Code => {
                let text = entry.content.clone();
                if let Err(error) = self.platform.write_text_to_clipboard(&text) {
                    self.last_status = format!("Clipboard write failed: {error}");
                    return;
                }
                compact(&text, 40)
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
        match self.platform.read_file_paths_from_clipboard() {
            Ok(Some(paths)) => {
                self.capture_clipboard_files(paths, cx);
                return;
            }
            Ok(None) => {}
            Err(error) => {
                self.last_status = format!("Clipboard file read failed: {error}");
            }
        }

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

    fn capture_clipboard_files(&mut self, file_paths: Vec<String>, cx: &mut Context<Self>) {
        let file_paths = file_paths
            .into_iter()
            .map(|path| path.trim().to_string())
            .filter(|path| !path.is_empty())
            .collect::<Vec<_>>();
        if file_paths.is_empty() {
            self.last_status = "Clipboard update ignored; empty file list".to_string();
            return;
        }

        let mut incoming = entry_from_clipboard_files(file_paths, OffsetDateTime::now_utc());
        incoming.source_app = self.current_source_app_name();

        if let Some(existing) = self.document.entries.first_mut() {
            if is_likely_duplicate(existing, &incoming, Duration::from_secs(2)) {
                *existing = paste_core::merge_entries(existing, &incoming);
                self.last_status = "Merged duplicate clipboard files".to_string();
                self.normalize_document_for_save();
                self.schedule_history_save(cx);
                return;
            }
        }

        self.document.entries.insert(0, incoming);
        self.normalize_document_for_save();
        self.selected_index = 0;
        self.pending_delete = None;
        self.start_selection_animation(cx);
        self.last_status = format!(
            "Captured file item; {} entries",
            self.document.entries.len()
        );
        self.schedule_history_save(cx);
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
        self.pending_delete = None;
        self.start_selection_animation(cx);
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
        self.pending_delete = None;
        self.start_selection_animation(cx);
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
        match &self.active_filter {
            OverlayFilter::All => paste_core::visible_entries(
                &self.document.entries,
                self.is_trash_view,
                None,
                &self.search_query,
            ),
            OverlayFilter::Favorites => visible_favorite_entries(
                &self.document.entries,
                self.is_trash_view,
                &self.search_query,
            ),
            OverlayFilter::Kind(kind) => visible_kind_entries(
                &self.document.entries,
                self.is_trash_view,
                kind.clone(),
                &self.search_query,
            ),
            OverlayFilter::Group(group_id) => paste_core::visible_entries(
                &self.document.entries,
                self.is_trash_view,
                Some(group_id),
                &self.search_query,
            ),
        }
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
    ToggleFavorite,
    MoveToTrash,
    RestoreFromTrash,
    PinQuick,
    PinWork,
    PinIdea,
    Unpin,
    ShowAll,
    ShowFavorites,
    FilterText,
    FilterLink,
    FilterImage,
    FilterFile,
    FilterCode,
    ShowQuick,
    ShowWork,
    ShowIdea,
    ToggleContextMenu,
    ToggleSettings,
    HideOverlay,
}

fn command_action_changes_selection_surface(action: CommandPaletteAction) -> bool {
    matches!(
        action,
        CommandPaletteAction::ToggleSearch
            | CommandPaletteAction::ToggleTrash
            | CommandPaletteAction::ShowAll
            | CommandPaletteAction::ShowFavorites
            | CommandPaletteAction::FilterText
            | CommandPaletteAction::FilterLink
            | CommandPaletteAction::FilterImage
            | CommandPaletteAction::FilterFile
            | CommandPaletteAction::FilterCode
            | CommandPaletteAction::ShowQuick
            | CommandPaletteAction::ShowWork
            | CommandPaletteAction::ShowIdea
    )
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
        CommandPaletteCommand {
            label: "Toggle favorite",
            hint: "Pinboard",
            action: CommandPaletteAction::ToggleFavorite,
        },
        trash_command,
        CommandPaletteCommand {
            label: "Show all",
            hint: "Tab",
            action: CommandPaletteAction::ShowAll,
        },
        CommandPaletteCommand {
            label: "Show favorites",
            hint: "Tab",
            action: CommandPaletteAction::ShowFavorites,
        },
        CommandPaletteCommand {
            label: "Filter Text",
            hint: "Category",
            action: CommandPaletteAction::FilterText,
        },
        CommandPaletteCommand {
            label: "Filter Links",
            hint: "Category",
            action: CommandPaletteAction::FilterLink,
        },
        CommandPaletteCommand {
            label: "Filter Images",
            hint: "Category",
            action: CommandPaletteAction::FilterImage,
        },
        CommandPaletteCommand {
            label: "Filter Files",
            hint: "Category",
            action: CommandPaletteAction::FilterFile,
        },
        CommandPaletteCommand {
            label: "Filter Code",
            hint: "Category",
            action: CommandPaletteAction::FilterCode,
        },
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
            label: "Show Quick pinboard",
            hint: "Pinboard",
            action: CommandPaletteAction::ShowQuick,
        },
        CommandPaletteCommand {
            label: "Show Work pinboard",
            hint: "Pinboard",
            action: CommandPaletteAction::ShowWork,
        },
        CommandPaletteCommand {
            label: "Show Idea pinboard",
            hint: "Pinboard",
            action: CommandPaletteAction::ShowIdea,
        },
        CommandPaletteCommand {
            label: "Open context menu",
            hint: "Actions",
            action: CommandPaletteAction::ToggleContextMenu,
        },
        CommandPaletteCommand {
            label: "Open settings",
            hint: "Panel",
            action: CommandPaletteAction::ToggleSettings,
        },
        CommandPaletteCommand {
            label: "Hide overlay",
            hint: "Esc",
            action: CommandPaletteAction::HideOverlay,
        },
    ]
}

fn render_entry_card(
    entry: &ClipboardEntry,
    selected: bool,
    selection_pulse: f32,
    reduced_motion: bool,
) -> impl IntoElement {
    let theme = app_theme();
    let accent = kind_accent(&entry.kind);
    let content_is_dark = matches!(entry.kind, ClipboardKind::Code);
    let image_card = matches!(entry.kind, ClipboardKind::Image);
    let dark_card = content_is_dark || image_card;
    let selected_border = THEME_SELECTION_BLUE;
    let body_bg = match entry.kind {
        ClipboardKind::Text => rgb(0xFFF3BF),
        ClipboardKind::Image => rgb(THEME_BLUE),
        ClipboardKind::Code => rgb(0x111827),
        ClipboardKind::File | ClipboardKind::Link => rgb(0xFFFDF7),
    };
    let hover_bg = match entry.kind {
        ClipboardKind::Text => rgb(0xFFECA3),
        ClipboardKind::Image => rgb(0x368FE0),
        ClipboardKind::Code => rgb(0x151E2D),
        ClipboardKind::File | ClipboardKind::Link => rgb(theme.card_hover),
    };
    let body_text: gpui::Hsla = if dark_card {
        rgb(0xFFFFFF).into()
    } else {
        rgb(theme.card_text).into()
    };
    let meta_text = if dark_card {
        rgb(0xE8F3FF)
    } else {
        rgb(theme.weak)
    };
    let icon_bg = if dark_card {
        alpha_rgb(0xFFFFFF, 0x32)
    } else {
        alpha_rgb(accent, 0x24)
    };
    let icon_text = if dark_card {
        rgb(0xFFFFFF)
    } else {
        rgb(accent)
    };
    let code_bg = if content_is_dark {
        rgb(0x111827)
    } else {
        body_bg
    };

    div()
        .id(format!(
            "entry-{}-{}",
            entry.kind_label(),
            entry.copied_at.unix_timestamp_nanos()
        ))
        .flex()
        .flex_col()
        .flex_none()
        .w(px(RAIL_CARD_WIDTH))
        .h(px(RAIL_CARD_HEIGHT))
        .relative()
        .rounded(px(CARD_RADIUS))
        .overflow_hidden()
        .border_1()
        .border_color(if selected {
            rgb(selected_border)
        } else {
            rgb(theme.border)
        })
        .bg(code_bg)
        .opacity(if selected && !reduced_motion {
            0.97 + (selection_pulse * 0.03)
        } else {
            1.0
        })
        .shadow_lg()
        .hover(move |style| style.bg(hover_bg))
        .text_color(body_text)
        .text_size(px(12.0))
        .when(selected, |card| card.border_2().shadow_2xl())
        .child(
            div()
                .flex()
                .justify_between()
                .items_center()
                .px(px(14.0))
                .pt(px(12.0))
                .text_color(body_text)
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(16.0))
                                .line_height(px(20.0))
                                .child(entry.kind_label()),
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .line_height(px(16.0))
                                .text_color(meta_text)
                                .child(copied_at_label(entry.copied_at)),
                        ),
                )
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_center()
                        .w(px(22.0))
                        .h(px(22.0))
                        .rounded(px(7.0))
                        .bg(icon_bg)
                        .text_color(icon_text)
                        .text_size(px(11.0))
                        .child(kind_badge_label(&entry.kind)),
                ),
        )
        .child(
            div()
                .flex()
                .flex_col()
                .flex_1()
                .justify_between()
                .px(px(14.0))
                .pt(px(9.0))
                .pb(px(12.0))
                .child(render_card_preview(entry, body_text))
                .child(render_card_footer(entry, accent, body_text)),
        )
        .when(selected, |card| {
            card.child(
                div()
                    .absolute()
                    .top_0()
                    .left_0()
                    .right_0()
                    .h(px(58.0))
                    .bg(alpha_rgb(0xFFFFFF, 0x16)),
            )
        })
}

fn render_card_preview(entry: &ClipboardEntry, body_text: gpui::Hsla) -> impl IntoElement {
    let theme = app_theme();
    match entry.kind {
        ClipboardKind::Link => div()
            .flex()
            .flex_col()
            .gap_2()
            .child(
                div()
                    .flex()
                    .items_center()
                    .gap_2()
                    .h(px(58.0))
                    .rounded_lg()
                    .bg(alpha_rgb(kind_accent(&ClipboardKind::Link), 0x24))
                    .border_1()
                    .border_color(alpha_rgb(kind_accent(&ClipboardKind::Link), 0x38))
                    .px(px(10.0))
                    .child(
                        div()
                            .w(px(24.0))
                            .h(px(24.0))
                            .rounded_full()
                            .bg(alpha_rgb(kind_accent(&ClipboardKind::Link), 0x48)),
                    )
                    .child(
                        div()
                            .flex_1()
                            .h(px(30.0))
                            .rounded(px(8.0))
                            .border_1()
                            .border_color(alpha_rgb(kind_accent(&ClipboardKind::Link), 0x2D))
                            .bg(alpha_rgb(0xFFFFFF, 0x18)),
                    ),
            )
            .child(
                div()
                    .text_size(px(14.0))
                    .line_height(px(17.0))
                    .text_color(body_text)
                    .child(compact(&entry_title(entry), 42)),
            )
            .child(
                div()
                    .text_size(px(12.0))
                    .line_height(px(15.0))
                    .text_color(rgb(theme.weak))
                    .child(compact(&entry_body(entry), 46)),
            ),
        ClipboardKind::Image => div()
            .flex()
            .flex_col()
            .gap_2()
            .child(
                div()
                    .h(px(96.0))
                    .rounded_lg()
                    .bg(alpha_rgb(kind_accent(&ClipboardKind::Image), 0x30))
                    .border_1()
                    .border_color(alpha_rgb(kind_accent(&ClipboardKind::Image), 0x72))
                    .flex()
                    .items_center()
                    .justify_center()
                    .text_color(rgb(0xFFFFFF))
                    .text_size(px(13.0))
                    .child("Image preview"),
            )
            .child(
                div()
                    .text_size(px(12.0))
                    .text_color(rgb(theme.weak))
                    .child("Copied image"),
            ),
        ClipboardKind::File => div()
            .flex()
            .flex_col()
            .gap_2()
            .child(
                div()
                    .flex()
                    .items_center()
                    .justify_center()
                    .h(px(78.0))
                    .rounded_lg()
                    .bg(alpha_rgb(file_primary_color(entry), 0x24))
                    .border_1()
                    .border_color(alpha_rgb(file_primary_color(entry), 0x42))
                    .child(
                        div()
                            .px(px(14.0))
                            .py(px(10.0))
                            .rounded(px(10.0))
                            .bg(alpha_rgb(file_primary_color(entry), 0x38))
                            .text_color(rgb(file_primary_color(entry)))
                            .text_size(px(18.0))
                            .child(file_primary_extension(entry)),
                    ),
            )
            .child(
                div()
                    .text_size(px(14.0))
                    .line_height(px(17.0))
                    .text_color(body_text)
                    .child(compact(&file_entry_label(entry), 32)),
            )
            .child(render_file_tokens(entry)),
        ClipboardKind::Code => div()
            .flex()
            .flex_col()
            .gap_2()
            .child(
                div()
                    .text_size(px(14.0))
                    .line_height(px(18.0))
                    .child(entry_title(entry)),
            )
            .child(
                div()
                    .h(px(104.0))
                    .overflow_hidden()
                    .p(px(8.0))
                    .rounded_lg()
                    .bg(rgb(0x0A0910))
                    .border_1()
                    .border_color(rgb(0x2D2835))
                    .font_family("Consolas")
                    .text_size(px(12.0))
                    .line_height(px(15.0))
                    .child(numbered_code_preview(&entry.content)),
            ),
        ClipboardKind::Text => div()
            .flex()
            .flex_col()
            .gap_2()
            .child(
                div()
                    .text_size(px(14.0))
                    .line_height(px(18.0))
                    .text_color(body_text)
                    .child(compact(&entry_title(entry), 54)),
            )
            .child(
                div()
                    .text_size(px(13.0))
                    .line_height(px(18.0))
                    .text_color(rgb(theme.muted))
                    .child(compact(&entry_body(entry), 118)),
            ),
    }
}

fn numbered_code_preview(content: &str) -> String {
    let mut lines = content
        .lines()
        .take(4)
        .enumerate()
        .map(|(index, line)| format!("{:>2}  {}", index + 1, compact(line, 20)));
    let preview = lines.by_ref().collect::<Vec<_>>().join("\n");
    if preview.is_empty() {
        compact(content, 96)
    } else {
        preview
    }
}

fn file_primary_extension(entry: &ClipboardEntry) -> String {
    entry
        .file_extensions
        .first()
        .map(|extension| extension.to_ascii_uppercase())
        .unwrap_or_else(|| "FILE".to_string())
}

fn file_primary_color(entry: &ClipboardEntry) -> u32 {
    entry
        .file_extensions
        .first()
        .map(|extension| extension_color(extension))
        .unwrap_or(0xA1A1AA)
}

fn extension_color(extension: &str) -> u32 {
    match extension
        .trim_start_matches('.')
        .to_ascii_lowercase()
        .as_str()
    {
        "pdf" => THEME_ERROR,
        "doc" | "docx" => THEME_BLUE,
        "xls" | "xlsx" => THEME_SUCCESS,
        "ppt" | "pptx" => THEME_ORANGE,
        "zip" | "7z" | "rar" => THEME_WARNING,
        _ => 0xA1A1AA,
    }
}

fn pinboard_color(group_id: &str) -> u32 {
    match group_id {
        WORK_GROUP_ID => THEME_PURPLE,
        QUICK_GROUP_ID => THEME_SUCCESS,
        IDEA_GROUP_ID => THEME_WARNING,
        "design" => THEME_ORANGE,
        "code" => THEME_BLUE,
        "archive" => 0xA1A1AA,
        _ => THEME_PINBOARD,
    }
}

fn render_file_tokens(entry: &ClipboardEntry) -> impl IntoElement {
    entry
        .file_extensions
        .iter()
        .take(3)
        .fold(div().flex().gap_1(), |row, extension| {
            let color = extension_color(extension);
            row.child(
                div()
                    .px(px(8.0))
                    .py(px(4.0))
                    .rounded_full()
                    .bg(alpha_rgb(color, 0x28))
                    .text_color(rgb(color))
                    .text_size(px(12.0))
                    .child(extension.to_ascii_uppercase()),
            )
        })
}

fn render_card_footer(
    entry: &ClipboardEntry,
    _accent: u32,
    body_text: gpui::Hsla,
) -> impl IntoElement {
    let theme = app_theme();
    div()
        .flex()
        .justify_between()
        .items_center()
        .text_size(px(12.0))
        .text_color(body_text.opacity(0.62))
        .child(compact(&entry.source_app, 18))
        .child(if entry.is_favorite {
            "Favorite".to_string()
        } else {
            entry
                .code_language
                .clone()
                .or_else(|| entry.file_extensions.first().cloned())
                .unwrap_or_else(|| kind_label(&entry.kind).to_ascii_lowercase())
        })
        .when(entry.is_favorite, |footer| {
            footer.child(
                div()
                    .absolute()
                    .right(px(12.0))
                    .bottom(px(12.0))
                    .px(px(8.0))
                    .py(px(4.0))
                    .rounded_full()
                    .bg(alpha_rgb(THEME_WARNING, 0x30))
                    .text_color(rgb(if matches!(entry.kind, ClipboardKind::Code) {
                        THEME_WARNING
                    } else {
                        theme.card_text
                    }))
                    .child("Star"),
            )
        })
}

fn render_pinboard_toolbar(groups: &[paste_core::PinnedGroup]) -> impl IntoElement {
    let theme = app_theme();
    let base = div()
        .flex()
        .items_center()
        .justify_center()
        .gap(px(26.0))
        .h(px(48.0))
        .child(
            div()
                .w(px(22.0))
                .h(px(22.0))
                .relative()
                .rounded_full()
                .border_2()
                .border_color(rgb(0x2C2018))
                .child(
                    div()
                        .absolute()
                        .right(px(-2.0))
                        .bottom(px(1.0))
                        .w(px(8.0))
                        .h(px(2.0))
                        .rounded_full()
                        .bg(rgb(0x2C2018)),
                ),
        )
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .px(px(16.0))
                .py(px(8.0))
                .rounded_full()
                .bg(alpha_rgb(theme.surface, 0x9D))
                .shadow_sm()
                .child(
                    div()
                        .w(px(17.0))
                        .h(px(17.0))
                        .rounded_full()
                        .border_2()
                        .border_color(rgb(theme.weak)),
                )
                .child(
                    div()
                        .text_size(px(16.0))
                        .text_color(rgb(theme.text))
                        .child("Clipboard"),
                ),
        );

    groups
        .iter()
        .take(4)
        .enumerate()
        .fold(base, |toolbar, (index, group)| {
            let fallback_color = match index {
                0 => 0xFF3B30,
                1 => THEME_WARNING,
                2 => THEME_SUCCESS,
                _ => THEME_BLUE,
            };
            let color = if group.id == QUICK_GROUP_ID
                || group.id == WORK_GROUP_ID
                || group.id == IDEA_GROUP_ID
            {
                pinboard_color(&group.id)
            } else {
                fallback_color
            };
            toolbar.child(
                div()
                    .flex()
                    .items_center()
                    .gap_2()
                    .text_size(px(16.0))
                    .text_color(rgb(theme.text))
                    .child(div().w(px(14.0)).h(px(14.0)).rounded_full().bg(rgb(color)))
                    .child(group.name.clone()),
            )
        })
}

fn render_preview(
    entry: Option<&ClipboardEntry>,
    selection_pulse: f32,
    reduced_motion: bool,
) -> impl IntoElement {
    let theme = app_theme();
    match entry {
        Some(entry) => {
            let accent = kind_accent(&entry.kind);
            let pulse_alpha = if reduced_motion {
                0x34
            } else {
                0x34 + ((selection_pulse * 0x24 as f32).round() as u32)
            };
            div()
                .flex()
                .justify_center()
                .h(px(PREVIEW_POPOVER_HEIGHT))
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap_2()
                        .w(px(640.0))
                        .h(px(PREVIEW_POPOVER_HEIGHT))
                        .px(px(20.0))
                        .py(px(14.0))
                        .rounded(px(24.0))
                        .border_1()
                        .border_color(alpha_rgb(accent, pulse_alpha))
                        .bg(rgb(theme.panel_alt))
                        .shadow_2xl()
                        .child(
                            div()
                                .flex()
                                .items_center()
                                .justify_between()
                                .child(div().text_size(px(14.0)).text_color(rgb(theme.text)).child(
                                    format!(
                                        "{} from {}",
                                        entry.kind_label(),
                                        compact(&entry.source_app, 28)
                                    ),
                                ))
                                .child(
                                    div()
                                        .px(px(10.0))
                                        .py(px(4.0))
                                        .rounded_full()
                                        .bg(alpha_rgb(accent, 0x34))
                                        .text_color(rgb(accent))
                                        .text_size(px(12.0))
                                        .child(copied_at_label(entry.copied_at)),
                                ),
                        )
                        .child(
                            div()
                                .flex()
                                .gap_2()
                                .child(div().w(px(4.0)).h(px(34.0)).rounded_full().bg(rgb(accent)))
                                .child(
                                    div()
                                        .flex_1()
                                        .text_size(px(15.0))
                                        .line_height(px(20.0))
                                        .text_color(rgb(theme.muted))
                                        .child(compact(&entry_body(entry), 155)),
                                ),
                        ),
                )
        }
        None => div().h(px(0.0)),
    }
}

fn render_delete_confirmation(entry: Option<&ClipboardEntry>) -> impl IntoElement {
    let theme = app_theme();
    div()
        .flex()
        .flex_col()
        .gap_2()
        .w(px(320.0))
        .p(px(16.0))
        .rounded(px(18.0))
        .border_1()
        .border_color(alpha_rgb(THEME_ERROR, 0x58))
        .bg(rgb(theme.panel_alt))
        .text_color(rgb(theme.text))
        .shadow_2xl()
        .child(div().text_size(px(15.0)).child("Delete this item?"))
        .child(
            div()
                .text_size(px(12.0))
                .text_color(rgb(theme.muted))
                .child(format!(
                    "{} will be moved to trash.",
                    entry
                        .map(entry_title)
                        .unwrap_or_else(|| "The selected item".to_string())
                )),
        )
        .child(
            div()
                .flex()
                .justify_end()
                .gap_2()
                .child(
                    div()
                        .px(px(12.0))
                        .py(px(7.0))
                        .rounded_full()
                        .bg(rgb(theme.surface))
                        .text_color(rgb(theme.muted))
                        .child("Esc Cancel"),
                )
                .child(
                    div()
                        .px(px(12.0))
                        .py(px(7.0))
                        .rounded_full()
                        .bg(rgb(THEME_ERROR))
                        .text_color(rgb(0xFFFFFF))
                        .child("Enter Delete"),
                ),
        )
}

fn render_context_menu(entry: Option<&ClipboardEntry>) -> impl IntoElement {
    let theme = app_theme();
    let pinboard = entry
        .and_then(|entry| entry.pinned_group_id.as_deref())
        .map(|group| format!("Pinboard: {group}"))
        .unwrap_or_else(|| "Pin".to_string());
    [
        ("Open", "Enter"),
        ("Paste to target", ""),
        ("Paste as Plain Text", "Shift Enter"),
        ("Copy", "Ctrl C"),
        ("Edit", "Ctrl E"),
        ("Rename", "Ctrl R"),
        ("Delete", "Del"),
        (pinboard.as_str(), ""),
        ("Quick Look", "Space"),
        ("Share", ""),
    ]
    .into_iter()
    .fold(
        div()
            .flex()
            .flex_col()
            .w(px(300.0))
            .p(px(8.0))
            .rounded(px(18.0))
            .border_1()
            .border_color(rgb(theme.border))
            .bg(alpha_rgb(theme.panel_alt, 0xEE))
            .text_size(px(14.0))
            .shadow_2xl(),
        |menu, (label, shortcut)| {
            let highlighted = label == "Paste as Plain Text";
            menu.child(
                div()
                    .h(px(28.0))
                    .flex()
                    .items_center()
                    .justify_between()
                    .px(px(10.0))
                    .rounded(px(10.0))
                    .bg(if highlighted {
                        rgb(THEME_SELECTION_BLUE)
                    } else {
                        rgba(0x00000000)
                    })
                    .text_color(if highlighted {
                        rgb(0xFFFFFF)
                    } else if label == "Delete" {
                        rgb(THEME_ERROR)
                    } else {
                        rgb(theme.text)
                    })
                    .hover(move |style| style.bg(rgb(theme.surface)))
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .gap_2()
                            .child(
                                div()
                                    .w(px(16.0))
                                    .h(px(16.0))
                                    .rounded(px(4.0))
                                    .border_1()
                                    .border_color(if highlighted {
                                        rgb(0xFFFFFF)
                                    } else if label == "Delete" {
                                        rgb(THEME_ERROR)
                                    } else {
                                        rgb(theme.weak)
                                    }),
                            )
                            .child(label.to_string()),
                    )
                    .child(
                        div()
                            .text_size(px(12.0))
                            .text_color(if highlighted {
                                alpha_rgb(0xFFFFFF, 0xCC)
                            } else {
                                rgb(theme.weak)
                            })
                            .child(shortcut.to_string()),
                    ),
            )
        },
    )
}

fn render_settings_panel(
    is_plain_text: bool,
    is_trash_view: bool,
    filter: &OverlayFilter,
) -> impl IntoElement {
    let theme = app_theme();
    let sections = [
        "Appearance",
        "Cards",
        "Shortcuts",
        "Pinboards",
        "Privacy",
        "About",
    ];
    let nav = sections.into_iter().fold(
        div().flex().flex_col().gap_1().w(px(128.0)),
        |nav, section| {
            nav.child(
                div()
                    .px(px(10.0))
                    .py(px(8.0))
                    .rounded(px(10.0))
                    .bg(if section == "Appearance" {
                        alpha_rgb(THEME_PURPLE, 0x30)
                    } else {
                        rgb(theme.panel_alt)
                    })
                    .text_color(if section == "Appearance" {
                        rgb(THEME_PURPLE)
                    } else {
                        rgb(theme.muted)
                    })
                    .child(section),
            )
        },
    );

    div()
        .flex()
        .gap_4()
        .w(px(560.0))
        .p(px(18.0))
        .rounded(px(24.0))
        .border_1()
        .border_color(rgb(theme.border))
        .bg(rgb(theme.panel_alt))
        .text_size(px(13.0))
        .shadow_2xl()
        .child(nav)
        .child(
            div()
                .flex()
                .flex_col()
                .gap_3()
                .flex_1()
                .child(
                    div()
                        .text_size(px(18.0))
                        .text_color(rgb(theme.text))
                        .child("Appearance"),
                )
                .child(render_setting_row(
                    "Theme",
                    "Dark / Light via PASTE_GPUI_THEME",
                ))
                .child(render_setting_row(
                    "Paste mode",
                    if is_plain_text {
                        "Plain text"
                    } else {
                        "Rich content"
                    },
                ))
                .child(render_setting_row(
                    "View",
                    if is_trash_view { "Trash" } else { "History" },
                ))
                .child(render_setting_row("Filter", &filter.label()))
                .child(render_setting_row(
                    "Motion",
                    "Uses Windows animation setting",
                )),
        )
}

fn render_setting_row(label: &str, value: &str) -> impl IntoElement {
    let theme = app_theme();
    div()
        .flex()
        .items_center()
        .justify_between()
        .h(px(38.0))
        .px(px(12.0))
        .rounded(px(10.0))
        .bg(rgb(theme.surface))
        .child(div().text_color(rgb(theme.text)).child(label.to_string()))
        .child(
            div()
                .px(px(10.0))
                .py(px(5.0))
                .rounded_full()
                .bg(rgb(theme.key_bg))
                .text_color(rgb(theme.muted))
                .child(value.to_string()),
        )
}

fn render_empty_state() -> impl IntoElement {
    let theme = app_theme();
    div()
        .flex()
        .flex_col()
        .items_center()
        .justify_center()
        .gap_3()
        .h(px(RAIL_CARD_HEIGHT))
        .child(
            div()
                .w(px(56.0))
                .h(px(56.0))
                .rounded(px(18.0))
                .border_2()
                .border_color(alpha_rgb(THEME_PURPLE, 0x73))
                .bg(alpha_rgb(THEME_PURPLE, 0x22)),
        )
        .child(
            div()
                .text_size(px(16.0))
                .text_color(rgb(theme.text))
                .child("No clipboard items yet"),
        )
        .child(
            div()
                .text_size(px(13.0))
                .text_color(rgb(theme.muted))
                .child("Copy text, images, links or files to see them here."),
        )
}

fn render_skeleton_rail() -> impl IntoElement {
    let theme = app_theme();
    (0..5).fold(
        div().flex().gap(px(RAIL_CARD_GAP)).h(px(RAIL_CARD_HEIGHT)),
        |rail, index| {
            let width = 88.0 + (index as f32 * 9.0);
            rail.child(
                div()
                    .flex()
                    .flex_col()
                    .gap_3()
                    .w(px(RAIL_CARD_WIDTH))
                    .h(px(RAIL_CARD_HEIGHT))
                    .p(px(14.0))
                    .rounded(px(CARD_RADIUS))
                    .border_1()
                    .border_color(rgb(theme.border))
                    .bg(rgb(theme.card))
                    .child(
                        div()
                            .h(px(22.0))
                            .w(px(width))
                            .rounded_full()
                            .bg(rgb(theme.surface)),
                    )
                    .child(
                        div()
                            .h(px(96.0))
                            .rounded_lg()
                            .bg(alpha_rgb(theme.surface, 0xCC)),
                    )
                    .child(
                        div()
                            .h(px(14.0))
                            .w(px(width + 34.0))
                            .rounded_full()
                            .bg(rgb(theme.surface)),
                    ),
            )
        },
    )
}

trait EntryKindLabel {
    fn kind_label(&self) -> &'static str;
}

impl EntryKindLabel for ClipboardEntry {
    fn kind_label(&self) -> &'static str {
        kind_label(&self.kind)
    }
}

fn kind_label(kind: &ClipboardKind) -> &'static str {
    match kind {
        ClipboardKind::Text => "Text",
        ClipboardKind::Image => "Image",
        ClipboardKind::Link => "Link",
        ClipboardKind::File => "File",
        ClipboardKind::Code => "Code",
    }
}

fn kind_badge_label(kind: &ClipboardKind) -> &'static str {
    match kind {
        ClipboardKind::Text => "T",
        ClipboardKind::Image => "I",
        ClipboardKind::Link => "L",
        ClipboardKind::File => "F",
        ClipboardKind::Code => "C",
    }
}

fn kind_accent(kind: &ClipboardKind) -> u32 {
    match kind {
        ClipboardKind::Text => 0x0A84FF,
        ClipboardKind::Image => THEME_BLUE,
        ClipboardKind::Link => THEME_BLUE,
        ClipboardKind::File => THEME_WARNING,
        ClipboardKind::Code => 0xA78BFA,
    }
}

fn app_theme() -> AppTheme {
    let dark_theme = env::var("PASTE_GPUI_THEME")
        .map(|value| value.trim().eq_ignore_ascii_case("dark"))
        .unwrap_or(false)
        || env_flag_enabled("PASTE_GPUI_DARK_THEME");

    if dark_theme {
        AppTheme {
            panel: 0x15151C,
            panel_alt: 0x202029,
            surface: 0x292936,
            card: 0x202029,
            card_hover: 0x292936,
            card_text: 0x18181B,
            text: 0xF4F4F5,
            muted: 0xA1A1AA,
            weak: 0x71717A,
            border: 0x2B2B36,
            key_bg: 0x2A2A34,
        }
    } else {
        AppTheme {
            panel: 0x2D2D2D,
            panel_alt: 0x3A3A3A,
            surface: 0x4A4A4A,
            card: 0xFFFFFF,
            card_hover: 0xFFF8EC,
            card_text: 0x241D18,
            text: 0xF4F4F5,
            muted: 0xD4D4D8,
            weak: 0xA1A1AA,
            border: 0x5A5A5A,
            key_bg: 0x454545,
        }
    }
}

fn render_panel_glass_layers(theme: AppTheme) -> impl IntoElement {
    div()
        .absolute()
        .top_0()
        .bottom_0()
        .left_0()
        .right_0()
        .rounded(px(PANEL_RADIUS))
        .overflow_hidden()
        .child(
            div()
                .absolute()
                .top_0()
                .left_0()
                .right_0()
                .h(px(112.0))
                .bg(linear_gradient(
                    180.0,
                    linear_color_stop(alpha_rgb(0xFFFFFF, PANEL_GLASS_HIGHLIGHT_ALPHA), 0.0),
                    linear_color_stop(alpha_rgb(0xFFFFFF, 0x00), 1.0),
                )),
        )
        .child(
            div()
                .absolute()
                .top_0()
                .bottom_0()
                .left_0()
                .w(px(220.0))
                .bg(linear_gradient(
                    90.0,
                    linear_color_stop(alpha_rgb(theme.surface, 0x24), 0.0),
                    linear_color_stop(alpha_rgb(theme.surface, 0x00), 1.0),
                )),
        )
        .child(
            div()
                .absolute()
                .bottom_0()
                .left_0()
                .right_0()
                .h(px(96.0))
                .bg(linear_gradient(
                    0.0,
                    linear_color_stop(alpha_rgb(0x000000, PANEL_GLASS_SHADE_ALPHA), 0.0),
                    linear_color_stop(alpha_rgb(0x000000, 0x00), 1.0),
                )),
        )
        .child(
            div()
                .absolute()
                .top_0()
                .left(px(24.0))
                .right(px(24.0))
                .h(px(1.0))
                .bg(alpha_rgb(0xFFFFFF, PANEL_GLASS_EDGE_ALPHA)),
        )
}

fn panel_glass_shadow() -> Vec<BoxShadow> {
    vec![
        BoxShadow {
            color: hsla(0.0, 0.0, 0.0, 0.28),
            offset: point(px(0.0), px(18.0)),
            blur_radius: px(48.0),
            spread_radius: px(-10.0),
            inset: false,
        },
        BoxShadow {
            color: hsla(0.0, 0.0, 1.0, 0.08),
            offset: point(px(0.0), px(1.0)),
            blur_radius: px(0.0),
            spread_radius: px(0.0),
            inset: true,
        },
        BoxShadow {
            color: hsla(0.0, 0.0, 0.0, 0.24),
            offset: point(px(0.0), px(-1.0)),
            blur_radius: px(0.0),
            spread_radius: px(0.0),
            inset: true,
        },
    ]
}

fn alpha_rgb(hex: u32, alpha: u32) -> gpui::Rgba {
    rgba((hex << 8) | alpha.min(0xFF))
}

fn selection_pulse_for_step(step: u64) -> f32 {
    let progress =
        (step.min(SELECTION_ANIMATION_STEPS) as f32) / (SELECTION_ANIMATION_STEPS as f32);
    (1.0 - progress).powi(2)
}

fn reduced_motion_requested(platform: &WindowsPlatformIntegration) -> bool {
    if env_flag_enabled("PASTE_GPUI_ENABLE_MOTION") {
        return false;
    }

    env_flag_enabled("PASTE_GPUI_REDUCED_MOTION") || !platform.client_area_animations_enabled()
}

fn env_flag_enabled(name: &str) -> bool {
    env::var(name)
        .map(|value| {
            matches!(
                value.trim().to_ascii_lowercase().as_str(),
                "1" | "true" | "yes" | "on"
            )
        })
        .unwrap_or(false)
}

fn entry_title(entry: &ClipboardEntry) -> String {
    match &entry.kind {
        ClipboardKind::Link => entry
            .link_title
            .clone()
            .or_else(|| entry.link_url.clone())
            .unwrap_or_else(|| compact(&entry.content, 60)),
        ClipboardKind::Image => "Image copied".to_string(),
        ClipboardKind::File => file_entry_label(entry),
        ClipboardKind::Code => entry
            .code_language
            .as_deref()
            .map(|language| format!("{} code", language))
            .unwrap_or_else(|| "Code snippet".to_string()),
        ClipboardKind::Text => compact(&entry.content, 60),
    }
}

fn entry_body(entry: &ClipboardEntry) -> String {
    match &entry.kind {
        ClipboardKind::File => {
            if entry.file_names.is_empty() {
                entry.file_paths.join(", ")
            } else {
                entry.file_names.join(", ")
            }
        }
        ClipboardKind::Image => entry.content.clone(),
        ClipboardKind::Link => entry
            .link_url
            .clone()
            .unwrap_or_else(|| compact(&entry.content, 90)),
        ClipboardKind::Code | ClipboardKind::Text => compact(&entry.content, 120),
    }
}

fn file_entry_label(entry: &ClipboardEntry) -> String {
    let count = entry.file_paths.len();
    match count {
        0 => "Files".to_string(),
        1 => entry
            .file_names
            .first()
            .cloned()
            .or_else(|| entry.file_paths.first().cloned())
            .unwrap_or_else(|| "File".to_string()),
        _ => format!("{count} files"),
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
        let selected_index = self.selected_index;
        let selected_entry = visible_entries.get(selected_index).cloned();
        let total_entries = visible_entries.len();
        let selection_pulse = self.selection_pulse;
        let visible_rail_cards = visible_rail_card_count(bottom_overlay_bounds(cx).size.width);
        let rail_start = if total_entries <= visible_rail_cards {
            0
        } else {
            selected_index
                .saturating_sub(visible_rail_cards / 2)
                .min(total_entries - visible_rail_cards)
        };
        let theme = app_theme();
        let rail = visible_entries
            .iter()
            .enumerate()
            .skip(rail_start)
            .take(visible_rail_cards)
            .fold(
                div()
                    .flex()
                    .gap(px(RAIL_CARD_GAP))
                    .h(px(RAIL_CARD_HEIGHT))
                    .overflow_hidden(),
                |rail, (index, entry)| {
                    rail.child(render_entry_card(
                        entry,
                        index == selected_index,
                        selection_pulse,
                        self.reduced_motion,
                    ))
                },
            );

        let rail_content = if env_flag_enabled("PASTE_GPUI_SHOW_SKELETON") {
            div()
                .relative()
                .h(px(RAIL_CARD_HEIGHT))
                .child(render_skeleton_rail())
        } else if total_entries == 0 {
            div().h(px(RAIL_CARD_HEIGHT)).child(render_empty_state())
        } else {
            div()
                .relative()
                .h(px(RAIL_CARD_HEIGHT))
                .overflow_hidden()
                .child(rail)
                .when(rail_start > 0, |rail| {
                    rail.child(
                        div()
                            .absolute()
                            .top_0()
                            .bottom_0()
                            .left_0()
                            .w(px(22.0))
                            .bg(alpha_rgb(theme.panel, 0xA8)),
                    )
                })
                .when(total_entries > rail_start + visible_rail_cards, |rail| {
                    rail.child(
                        div()
                            .absolute()
                            .top_0()
                            .bottom_0()
                            .right_0()
                            .w(px(22.0))
                            .bg(alpha_rgb(theme.panel, 0xA8)),
                    )
                })
        };

        let command_palette_commands = command_palette_commands(self.is_trash_view);
        let command_palette = command_palette_commands.iter().enumerate().fold(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .w(px(440.0))
                .p(px(12.0))
                .rounded(px(18.0))
                .border_1()
                .border_color(rgb(THEME_PURPLE))
                .text_size(px(13.0))
                .bg(rgb(theme.panel_alt))
                .shadow_2xl()
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
                            alpha_rgb(THEME_PURPLE, 0x38)
                        } else {
                            rgb(theme.panel_alt)
                        })
                        .text_color(rgb(theme.text))
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
            .on_action(cx.listener(Self::move_left))
            .on_action(cx.listener(Self::move_right))
            .on_action(cx.listener(Self::paste_selected))
            .on_action(cx.listener(Self::one_shot_plain_paste))
            .on_action(cx.listener(Self::close_overlay))
            .on_action(cx.listener(Self::toggle_trash))
            .on_action(cx.listener(Self::toggle_plain_text))
            .on_action(cx.listener(Self::toggle_command_palette))
            .on_action(cx.listener(Self::toggle_search))
            .on_action(cx.listener(Self::toggle_favorite))
            .on_action(cx.listener(Self::toggle_context_menu))
            .on_action(cx.listener(Self::toggle_settings))
            .on_action(cx.listener(Self::confirm_delete))
            .on_action(cx.listener(Self::cancel_delete))
            .on_action(cx.listener(Self::delete_selected))
            .on_action(cx.listener(Self::restore_selected))
            .on_action(cx.listener(Self::pin_quick))
            .on_action(cx.listener(Self::pin_work))
            .on_action(cx.listener(Self::pin_idea))
            .on_action(cx.listener(Self::unpin_selected))
            .relative()
            .flex()
            .flex_col()
            .justify_end()
            .gap_2()
            .size_full()
            .bg(rgba(0x00000000))
            .text_color(rgb(theme.text))
            .text_size(px(13.0))
            .when(env_flag_enabled("PASTE_GPUI_SHOW_PREVIEW"), |view| {
                view.child(render_preview(
                    selected_entry.as_ref(),
                    selection_pulse,
                    self.reduced_motion,
                ))
            })
            .child(
                div()
                    .relative()
                    .flex()
                    .flex_col()
                    .gap(px(18.0))
                    .h(px(MAIN_PANEL_HEIGHT))
                    .rounded(px(PANEL_RADIUS))
                    .overflow_hidden()
                    .bg(alpha_rgb(theme.panel, PANEL_BACKGROUND_ALPHA))
                    .shadow(panel_glass_shadow())
                    .px(px(PANEL_HORIZONTAL_PADDING))
                    .py(px(20.0))
                    .child(render_panel_glass_layers(theme))
                    .child(render_pinboard_toolbar(&self.document.groups))
                    .child(rail_content),
            )
            .when(self.pending_delete.is_some(), |view| {
                view.child(
                    div()
                        .absolute()
                        .right(px(24.0))
                        .bottom(px(74.0))
                        .child(render_delete_confirmation(selected_entry.as_ref())),
                )
            })
            .when(self.is_command_palette_open, |view| {
                view.child(
                    div()
                        .absolute()
                        .left_0()
                        .right_0()
                        .bottom(px(78.0))
                        .flex()
                        .justify_center()
                        .child(command_palette),
                )
            })
            .when(self.is_context_menu_open, |view| {
                view.child(
                    div()
                        .absolute()
                        .left_0()
                        .right_0()
                        .bottom(px(24.0))
                        .flex()
                        .justify_center()
                        .child(render_context_menu(selected_entry.as_ref())),
                )
            })
            .when(self.is_settings_open, |view| {
                view.child(div().absolute().right(px(24.0)).bottom(px(62.0)).child(
                    render_settings_panel(
                        self.is_plain_text,
                        self.is_trash_view,
                        &self.active_filter,
                    ),
                ))
            })
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
    let elapsed_seconds = (OffsetDateTime::now_utc() - copied_at)
        .whole_seconds()
        .max(0);
    if elapsed_seconds < 60 {
        return "now".to_string();
    }
    if elapsed_seconds < 60 * 60 {
        return format!("{}m", elapsed_seconds / 60);
    }
    if elapsed_seconds < 60 * 60 * 24 {
        return format!("{}h", elapsed_seconds / (60 * 60));
    }
    if elapsed_seconds < 60 * 60 * 24 * 7 {
        return format!("{}d", elapsed_seconds / (60 * 60 * 24));
    }

    copied_at
        .format(&time::macros::format_description!("[month]/[day]"))
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
            entry.kind = ClipboardKind::Link;
            entry.content = text;
            entry.link_url = Some(url.as_str().to_string());
            entry.link_title = None;
            return entry;
        }
    }

    if looks_like_code(trimmed) {
        let language = detect_code_language(trimmed);
        let mut entry = ClipboardEntry::text(text, copied_at);
        entry.kind = ClipboardKind::Code;
        entry.code_language = Some(language);
        return entry;
    }

    ClipboardEntry::text(text, copied_at)
}

fn entry_from_clipboard_files(
    file_paths: Vec<String>,
    copied_at: OffsetDateTime,
) -> ClipboardEntry {
    let file_names = file_paths
        .iter()
        .map(|path| {
            Path::new(path)
                .file_name()
                .and_then(|name| name.to_str())
                .unwrap_or(path)
                .to_string()
        })
        .collect::<Vec<_>>();
    let file_extensions = file_paths
        .iter()
        .filter_map(|path| {
            Path::new(path)
                .extension()
                .and_then(|extension| extension.to_str())
                .map(|extension| extension.to_lowercase())
        })
        .collect::<Vec<_>>();

    let mut entry = ClipboardEntry::text(format!("[Files: {}]", file_paths.len()), copied_at);
    entry.kind = ClipboardKind::File;
    entry.file_paths = file_paths;
    entry.file_names = file_names;
    entry.file_extensions = file_extensions;
    entry
}

fn entry_from_clipboard_image(
    png_bytes: Vec<u8>,
    dib_bytes: Vec<u8>,
    copied_at: OffsetDateTime,
) -> ClipboardEntry {
    let mut entry = ClipboardEntry::text(format!("[Image: {} bytes]", png_bytes.len()), copied_at);
    entry.kind = ClipboardKind::Image;
    entry.image_png_bytes = Some(png_bytes);
    entry.image_dib_bytes = Some(dib_bytes);
    entry
}

fn looks_like_code(text: &str) -> bool {
    if text.len() < 4 {
        return false;
    }

    if serde_json_like(text) {
        return true;
    }

    let code_markers = [
        "fn ",
        "let ",
        "use ",
        "impl ",
        "class ",
        "public ",
        "private ",
        "const ",
        "function ",
        "=>",
        "Write-Host",
        "param(",
        "using ",
        "#include",
    ];
    let marker_hits = code_markers
        .iter()
        .filter(|marker| text.contains(**marker))
        .count();
    let punctuation_hits = ['{', '}', ';', '(', ')']
        .iter()
        .filter(|ch| text.contains(**ch))
        .count();

    marker_hits >= 1 && (punctuation_hits >= 2 || text.lines().count() > 1)
}

fn detect_code_language(text: &str) -> String {
    let trimmed = text.trim();
    if serde_json_like(trimmed) {
        return "json".to_string();
    }
    if trimmed.contains("Write-Host") || trimmed.contains("param(") || trimmed.contains("$PS") {
        return "powershell".to_string();
    }
    if trimmed.contains("fn ") || trimmed.contains("impl ") || trimmed.contains("use ") {
        return "rust".to_string();
    }
    if trimmed.contains("using ")
        || trimmed.contains("namespace ")
        || trimmed.contains("public class")
    {
        return "csharp".to_string();
    }
    if trimmed.contains("function ") || trimmed.contains("=>") || trimmed.contains("console.") {
        return "javascript".to_string();
    }
    "plaintext".to_string()
}

fn serde_json_like(text: &str) -> bool {
    let trimmed = text.trim();
    ((trimmed.starts_with('{') && trimmed.ends_with('}'))
        || (trimmed.starts_with('[') && trimmed.ends_with(']')))
        && (trimmed.contains(':') || trimmed.contains(','))
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
        // Keep the OS window fully transparent; the panel itself fakes the frosted depth.
        window_background: WindowBackgroundAppearance::Transparent,
        window_decorations: Some(WindowDecorations::Client),
        window_bounds: Some(WindowBounds::Windowed(bottom_overlay_bounds(cx))),
        ..Default::default()
    }
}

fn bottom_overlay_bounds(cx: &App) -> Bounds<Pixels> {
    let Some(display) = cx.primary_display() else {
        return Bounds::centered(None, fallback_overlay_panel_size(), cx);
    };

    let display_bounds = display.bounds();
    let panel_size = overlay_panel_size_for_display(display_bounds);
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

fn overlay_panel_size_for_display(display_bounds: Bounds<Pixels>) -> gpui::Size<Pixels> {
    let desired_width = display_bounds.size.width * OVERLAY_SCREEN_WIDTH_RATIO;
    let safe_width = display_bounds.size.width - px(OVERLAY_SAFE_SIDE_MARGIN * 2.0);
    let panel_width = if safe_width > px(0.0) && desired_width > safe_width {
        safe_width
    } else {
        desired_width
    };

    size(panel_width, px(OVERLAY_WINDOW_HEIGHT))
}

fn visible_rail_card_count(panel_width: Pixels) -> usize {
    let rail_width = panel_width.as_f32() - (PANEL_HORIZONTAL_PADDING * 2.0);
    let card_stride = RAIL_CARD_WIDTH + RAIL_CARD_GAP;
    (((rail_width + RAIL_CARD_GAP) / card_stride).floor() as usize).max(1)
}

fn fallback_overlay_panel_size() -> gpui::Size<Pixels> {
    size(px(OVERLAY_FALLBACK_WIDTH), px(OVERLAY_WINDOW_HEIGHT))
}

fn main() {
    let (platform_loop, events, platform_status) = start_platform();

    application().run(|cx: &mut App| {
        cx.bind_keys([
            KeyBinding::new("up", MoveUp, Some("PasteOverlay")),
            KeyBinding::new("down", MoveDown, Some("PasteOverlay")),
            KeyBinding::new("left", MoveLeft, Some("PasteOverlay")),
            KeyBinding::new("right", MoveRight, Some("PasteOverlay")),
            KeyBinding::new("enter", PasteSelected, Some("PasteOverlay")),
            KeyBinding::new("alt-enter", OneShotPlainPaste, Some("PasteOverlay")),
            KeyBinding::new("escape", CloseOverlay, Some("PasteOverlay")),
            KeyBinding::new("ctrl-t", ToggleTrash, Some("PasteOverlay")),
            KeyBinding::new("ctrl-k", ToggleCommandPalette, Some("PasteOverlay")),
            KeyBinding::new("ctrl-f", ToggleSearch, Some("PasteOverlay")),
            KeyBinding::new("shift-f10", ToggleContextMenu, Some("PasteOverlay")),
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

        assert_eq!(entry.kind, ClipboardKind::Link);
        assert_eq!(entry.link_url.as_deref(), Some("https://example.com/path"));
    }

    #[test]
    fn clipboard_non_url_stays_text_entry() {
        let entry = entry_from_clipboard_text("not a url".to_string(), at(0));

        assert_eq!(entry.kind, ClipboardKind::Text);
        assert_eq!(entry.link_url, None);
    }

    #[test]
    fn clipboard_image_becomes_image_entry() {
        let entry = entry_from_clipboard_image(vec![1, 2, 3], vec![4, 5, 6], at(0));

        assert_eq!(entry.kind, ClipboardKind::Image);
        assert_eq!(entry.content, "[Image: 3 bytes]");
        assert_eq!(entry.image_png_bytes.as_deref(), Some(&[1, 2, 3][..]));
        assert_eq!(entry.image_dib_bytes.as_deref(), Some(&[4, 5, 6][..]));
    }

    #[test]
    fn clipboard_code_becomes_code_entry() {
        let entry =
            entry_from_clipboard_text("fn main() {\n    println!(\"hi\");\n}".to_string(), at(0));

        assert_eq!(entry.kind, ClipboardKind::Code);
        assert_eq!(entry.code_language.as_deref(), Some("rust"));
    }

    #[test]
    fn clipboard_files_become_file_entry() {
        let entry = entry_from_clipboard_files(
            vec![
                r"C:\Temp\report.pdf".to_string(),
                r"C:\Temp\image.png".to_string(),
            ],
            at(0),
        );

        assert_eq!(entry.kind, ClipboardKind::File);
        assert_eq!(entry.file_names, vec!["report.pdf", "image.png"]);
        assert_eq!(entry.file_extensions, vec!["pdf", "png"]);
    }

    #[test]
    fn selection_pulse_eases_down_to_zero() {
        assert!(selection_pulse_for_step(1) < 1.0);
        assert!(selection_pulse_for_step(2) < selection_pulse_for_step(1));
        assert_eq!(selection_pulse_for_step(SELECTION_ANIMATION_STEPS), 0.0);
    }

    #[test]
    fn overlay_panel_width_uses_90_percent_of_display() {
        let display_bounds = bounds(point(px(0.0), px(0.0)), size(px(1920.0), px(1080.0)));

        let panel_size = overlay_panel_size_for_display(display_bounds);

        assert_eq!(panel_size.width, px(1728.0));
        assert_eq!(panel_size.height, px(OVERLAY_WINDOW_HEIGHT));
    }

    #[test]
    fn overlay_panel_width_keeps_safe_side_margins_on_small_display() {
        let display_bounds = bounds(point(px(0.0), px(0.0)), size(px(360.0), px(720.0)));

        let panel_size = overlay_panel_size_for_display(display_bounds);

        assert_eq!(panel_size.width, px(296.0));
    }

    #[test]
    fn visible_rail_card_count_grows_with_panel_width() {
        assert_eq!(visible_rail_card_count(px(1120.0)), 5);
        assert_eq!(visible_rail_card_count(px(1728.0)), 8);
    }
}
