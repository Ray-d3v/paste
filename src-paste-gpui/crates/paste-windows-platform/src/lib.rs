use std::{
    mem::{size_of, size_of_val},
    sync::{
        atomic::{AtomicUsize, Ordering},
        mpsc, Mutex, OnceLock,
    },
    thread::{self, JoinHandle},
};
use thiserror::Error;

#[cfg(windows)]
use windows::{
    core::{w, Error as WindowsError, BOOL, PCWSTR, PWSTR},
    Win32::{
        Foundation::{CloseHandle, HANDLE, HINSTANCE, HWND, LPARAM, LRESULT, POINT, WPARAM},
        Graphics::Dwm::{
            DwmSetWindowAttribute, DWMWA_BORDER_COLOR, DWMWA_WINDOW_CORNER_PREFERENCE,
            DWMWCP_DONOTROUND,
        },
        Graphics::Gdi::{
            CreateRoundRectRgn, DeleteObject, GetDC, GetDIBits, GetObjectW, ReleaseDC,
            SetWindowRgn, BITMAP, BITMAPINFO, BITMAPINFOHEADER, BI_RGB, DIB_RGB_COLORS, HBITMAP,
            HGDIOBJ,
        },
        System::{
            DataExchange::{
                AddClipboardFormatListener, CloseClipboard, EmptyClipboard, GetClipboardData,
                IsClipboardFormatAvailable, OpenClipboard, RemoveClipboardFormatListener,
                SetClipboardData,
            },
            LibraryLoader::GetModuleHandleW,
            Memory::{
                GlobalAlloc, GlobalLock, GlobalSize, GlobalUnlock, GMEM_MOVEABLE, GMEM_ZEROINIT,
            },
            Ole::{CF_BITMAP, CF_DIB, CF_HDROP, CF_UNICODETEXT},
            Threading::{
                AttachThreadInput, GetCurrentProcessId, GetCurrentThreadId, OpenProcess,
                QueryFullProcessImageNameW, PROCESS_NAME_FORMAT, PROCESS_QUERY_LIMITED_INFORMATION,
            },
        },
        UI::{
            Input::KeyboardAndMouse::{
                RegisterHotKey, SendInput, SetActiveWindow, SetFocus, UnregisterHotKey,
                HOT_KEY_MODIFIERS, INPUT, INPUT_0, INPUT_KEYBOARD, KEYBDINPUT, KEYBD_EVENT_FLAGS,
                KEYEVENTF_KEYUP, MOD_ALT, MOD_CONTROL, VIRTUAL_KEY, VK_CONTROL, VK_INSERT,
                VK_SHIFT, VK_V,
            },
            Shell::{
                DragQueryFileW, Shell_NotifyIconW, DROPFILES, HDROP, NIF_ICON, NIF_MESSAGE,
                NIF_TIP, NIM_ADD, NIM_DELETE, NIM_SETVERSION, NOTIFYICONDATAW, NOTIFYICONDATAW_0,
                NOTIFYICON_VERSION_4, NOTIFY_ICON_DATA_FLAGS,
            },
            WindowsAndMessaging::{
                AllowSetForegroundWindow, AppendMenuW, BringWindowToTop, CreatePopupMenu,
                CreateWindowExW, DefWindowProcW, DestroyMenu, DestroyWindow, DispatchMessageW,
                EnumWindows, GetAncestor, GetCursorPos, GetForegroundWindow, GetGUIThreadInfo,
                GetMessageW, GetWindowTextLengthW, GetWindowTextW, GetWindowThreadProcessId,
                IsIconic, IsWindow, LoadIconW, PostMessageW, PostQuitMessage, RegisterClassW,
                SetForegroundWindow, SetWindowPos, ShowWindow, SystemParametersInfoW,
                TrackPopupMenuEx, TranslateMessage, GA_ROOT, GUITHREADINFO, HICON, HWND_TOPMOST,
                IDI_APPLICATION, MF_STRING, MSG, SPI_GETCLIENTAREAANIMATION, SWP_NOACTIVATE,
                SWP_NOOWNERZORDER, SW_HIDE, SW_RESTORE, SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS,
                TPM_BOTTOMALIGN, TPM_LEFTALIGN, TPM_RETURNCMD, TPM_RIGHTBUTTON, WINDOW_EX_STYLE,
                WINDOW_STYLE, WM_CLIPBOARDUPDATE, WM_CLOSE, WM_CONTEXTMENU, WM_DESTROY, WM_HOTKEY,
                WM_LBUTTONUP, WM_PASTE, WNDCLASSW,
            },
        },
    },
};

pub const OVERLAY_HOTKEY_ID: i32 = 0x5056;
pub const TRAY_ICON_ID: u32 = 0x5057;
pub const WM_TRAY_ICON: u32 = 0x8000 + 0x5057;
pub const WM_PASTE_GPUI_TEST_TRAY_COMMAND: u32 = 0x8000 + 0x5058;
pub const WM_PASTE_GPUI_TEST_PLATFORM_STATUS: u32 = 0x8000 + 0x5059;
pub const WM_PASTE_GPUI_TEST_TRAY_ERROR: u32 = 0x8000 + 0x505A;
pub const TEST_TRAY_COMMAND_SHOW: usize = 0x5201;
pub const TEST_TRAY_COMMAND_RESTART: usize = 0x5202;
pub const TEST_TRAY_COMMAND_EXIT: usize = 0x5203;
pub const PLATFORM_STATUS_HOTKEY: usize = 0x01;
pub const PLATFORM_STATUS_CLIPBOARD_LISTENER: usize = 0x02;
pub const PLATFORM_STATUS_TRAY_ICON: usize = 0x04;
const TRAY_MENU_SHOW: usize = 0x5101;
const TRAY_MENU_RESTART: usize = 0x5102;
const TRAY_MENU_EXIT: usize = 0x5103;
const DWMWA_COLOR_NONE: u32 = 0xFFFFFFFE;
const NIN_SELECT: u32 = 0x0400;
const NIN_KEYSELECT: u32 = 0x0401;
static PLATFORM_FEATURE_STATUS: AtomicUsize = AtomicUsize::new(0);
static PLATFORM_TRAY_LAST_ERROR: AtomicUsize = AtomicUsize::new(0);

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum PlatformEvent {
    OverlayHotkeyPressed,
    ClipboardUpdated,
    TrayShowRequested,
    TrayRestartRequested,
    TrayExitRequested,
    MessageLoopStopped,
}

pub struct PlatformEventLoop {
    window: WindowHandle,
    thread: Option<JoinHandle<()>>,
}

impl PlatformEventLoop {
    pub fn window(&self) -> WindowHandle {
        self.window
    }
}

impl Drop for PlatformEventLoop {
    fn drop(&mut self) {
        #[cfg(windows)]
        if !self.window.is_null() {
            unsafe {
                let _ = PostMessageW(Some(self.window.hwnd()), WM_CLOSE, WPARAM(0), LPARAM(0));
            }
        }

        if let Some(thread) = self.thread.take() {
            let _ = thread.join();
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct WindowHandle(pub isize);

impl WindowHandle {
    pub fn is_null(self) -> bool {
        self.0 == 0
    }

    #[cfg(windows)]
    fn hwnd(self) -> HWND {
        HWND(self.0 as *mut _)
    }

    #[cfg(windows)]
    fn from_hwnd(hwnd: HWND) -> Self {
        Self(hwnd.0 as isize)
    }
}

pub trait PlatformIntegration {
    fn register_overlay_hotkey(&self, window: WindowHandle) -> Result<(), PlatformError>;
    fn unregister_overlay_hotkey(&self, window: WindowHandle) -> Result<(), PlatformError>;
    fn add_clipboard_listener(&self, window: WindowHandle) -> Result<(), PlatformError>;
    fn remove_clipboard_listener(&self, window: WindowHandle) -> Result<(), PlatformError>;
    fn foreground_window(&self) -> Option<WindowHandle>;
    fn resolve_top_level_window(&self, window: WindowHandle) -> Option<WindowHandle>;
    fn window_process_name(&self, window: WindowHandle) -> Result<Option<String>, PlatformError>;
    fn is_window(&self, window: WindowHandle) -> bool;
    fn read_text_from_clipboard(&self) -> Result<Option<String>, PlatformError>;
    fn read_image_dib_from_clipboard(&self) -> Result<Option<Vec<u8>>, PlatformError>;
    fn read_file_paths_from_clipboard(&self) -> Result<Option<Vec<String>>, PlatformError>;
    fn write_text_to_clipboard(&self, text: &str) -> Result<(), PlatformError>;
    fn write_image_dib_to_clipboard(&self, dib_bytes: &[u8]) -> Result<(), PlatformError>;
    fn write_file_paths_to_clipboard(&self, file_paths: &[String]) -> Result<(), PlatformError>;
    fn paste_into_window(&self, target: WindowHandle) -> Result<(), PlatformError>;
    fn client_area_animations_enabled(&self) -> bool;
    fn find_window_by_title(&self, title: &str) -> Result<Option<WindowHandle>, PlatformError>;
    fn show_window(&self, window: WindowHandle) -> Result<(), PlatformError>;
    fn hide_window(&self, window: WindowHandle) -> Result<(), PlatformError>;
    fn move_window(
        &self,
        window: WindowHandle,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
    ) -> Result<(), PlatformError>;
    fn add_tray_icon(&self, window: WindowHandle) -> Result<(), PlatformError>;
    fn remove_tray_icon(&self, window: WindowHandle) -> Result<(), PlatformError>;
}

#[derive(Debug, Error)]
pub enum PlatformError {
    #[error("operation is not implemented")]
    NotImplemented,
    #[error("Windows API failed in {operation}: {source}")]
    Win32 {
        operation: &'static str,
        #[source]
        source: WindowsError,
    },
    #[error("invalid window handle")]
    InvalidWindow,
    #[error("clipboard data was not valid UTF-16")]
    InvalidClipboardText,
    #[error("platform event loop did not start")]
    EventLoopStartup,
}

#[derive(Default)]
pub struct WindowsPlatformIntegration;

impl WindowsPlatformIntegration {
    pub fn new() -> Self {
        Self
    }

    #[cfg(windows)]
    pub fn suppress_window_border(&self, window: WindowHandle) -> Result<(), PlatformError> {
        ensure_window(window)?;
        let color = DWMWA_COLOR_NONE;
        unsafe {
            DwmSetWindowAttribute(
                window.hwnd(),
                DWMWA_BORDER_COLOR,
                &color as *const _ as *const _,
                size_of::<u32>() as u32,
            )
        }
        .map_err(|source| PlatformError::Win32 {
            operation: "DwmSetWindowAttribute(DWMWA_BORDER_COLOR)",
            source,
        })
    }

    #[cfg(not(windows))]
    pub fn suppress_window_border(&self, _window: WindowHandle) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    #[cfg(windows)]
    pub fn apply_overlay_window_shape(
        &self,
        window: WindowHandle,
        width: i32,
        height: i32,
        radius: i32,
    ) -> Result<(), PlatformError> {
        ensure_window(window)?;

        let corner_preference = DWMWCP_DONOTROUND;
        let _ = unsafe {
            DwmSetWindowAttribute(
                window.hwnd(),
                DWMWA_WINDOW_CORNER_PREFERENCE,
                &corner_preference as *const _ as *const _,
                size_of_val(&corner_preference) as u32,
            )
        };

        let width = width.max(1);
        let height = height.max(1);
        let diameter = radius.max(1).saturating_mul(2);
        let region = unsafe { CreateRoundRectRgn(0, 0, width + 1, height + 1, diameter, diameter) };
        if region.is_invalid() {
            return Err(Self::win32_error("CreateRoundRectRgn"));
        }

        let result = unsafe { SetWindowRgn(window.hwnd(), Some(region), true) };
        if result == 0 {
            unsafe {
                let _ = DeleteObject(HGDIOBJ(region.0));
            }
            return Err(Self::win32_error("SetWindowRgn"));
        }

        Ok(())
    }

    #[cfg(not(windows))]
    pub fn apply_overlay_window_shape(
        &self,
        _window: WindowHandle,
        _width: i32,
        _height: i32,
        _radius: i32,
    ) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    #[cfg(windows)]
    fn win32_error(operation: &'static str) -> PlatformError {
        PlatformError::Win32 {
            operation,
            source: WindowsError::from_thread(),
        }
    }
}

pub fn spawn_platform_event_loop(
) -> Result<(PlatformEventLoop, mpsc::Receiver<PlatformEvent>), PlatformError> {
    #[cfg(windows)]
    {
        spawn_windows_event_loop()
    }

    #[cfg(not(windows))]
    {
        Err(PlatformError::NotImplemented)
    }
}

#[cfg(windows)]
impl PlatformIntegration for WindowsPlatformIntegration {
    fn register_overlay_hotkey(&self, window: WindowHandle) -> Result<(), PlatformError> {
        ensure_window(window)?;
        unsafe {
            RegisterHotKey(
                Some(window.hwnd()),
                OVERLAY_HOTKEY_ID,
                HOT_KEY_MODIFIERS(MOD_CONTROL.0 | MOD_ALT.0),
                VK_V.0 as u32,
            )
        }
        .map_err(|source| PlatformError::Win32 {
            operation: "RegisterHotKey",
            source,
        })
    }

    fn unregister_overlay_hotkey(&self, window: WindowHandle) -> Result<(), PlatformError> {
        ensure_window(window)?;
        unsafe { UnregisterHotKey(Some(window.hwnd()), OVERLAY_HOTKEY_ID) }.map_err(|source| {
            PlatformError::Win32 {
                operation: "UnregisterHotKey",
                source,
            }
        })
    }

    fn add_clipboard_listener(&self, window: WindowHandle) -> Result<(), PlatformError> {
        ensure_window(window)?;
        unsafe { AddClipboardFormatListener(window.hwnd()) }.map_err(|source| {
            PlatformError::Win32 {
                operation: "AddClipboardFormatListener",
                source,
            }
        })
    }

    fn remove_clipboard_listener(&self, window: WindowHandle) -> Result<(), PlatformError> {
        ensure_window(window)?;
        unsafe { RemoveClipboardFormatListener(window.hwnd()) }.map_err(|source| {
            PlatformError::Win32 {
                operation: "RemoveClipboardFormatListener",
                source,
            }
        })
    }

    fn foreground_window(&self) -> Option<WindowHandle> {
        let hwnd = unsafe { GetForegroundWindow() };
        (!hwnd.0.is_null()).then(|| WindowHandle::from_hwnd(hwnd))
    }

    fn resolve_top_level_window(&self, window: WindowHandle) -> Option<WindowHandle> {
        if window.is_null() {
            return None;
        }

        let root = unsafe { GetAncestor(window.hwnd(), GA_ROOT) };
        let resolved = if root.0.is_null() {
            window.hwnd()
        } else {
            root
        };
        unsafe { IsWindow(Some(resolved)).as_bool() }.then(|| WindowHandle::from_hwnd(resolved))
    }

    fn window_process_name(&self, window: WindowHandle) -> Result<Option<String>, PlatformError> {
        if window.is_null() {
            return Ok(None);
        }

        let mut process_id = 0;
        unsafe {
            GetWindowThreadProcessId(window.hwnd(), Some(&mut process_id));
        }
        if process_id == 0 {
            return Ok(None);
        }

        let process = unsafe { OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, process_id) }
            .map_err(|source| PlatformError::Win32 {
                operation: "OpenProcess",
                source,
            })?;

        let mut buffer = vec![0u16; 32_768];
        let mut len = buffer.len() as u32;
        let query = unsafe {
            QueryFullProcessImageNameW(
                process,
                PROCESS_NAME_FORMAT(0),
                PWSTR(buffer.as_mut_ptr()),
                &mut len,
            )
        };
        let _ = unsafe { CloseHandle(process) };

        query.map_err(|source| PlatformError::Win32 {
            operation: "QueryFullProcessImageNameW",
            source,
        })?;

        let path = String::from_utf16_lossy(&buffer[..len as usize]);
        Ok(process_name_from_path(&path))
    }

    fn is_window(&self, window: WindowHandle) -> bool {
        !window.is_null() && unsafe { IsWindow(Some(window.hwnd())).as_bool() }
    }

    fn read_text_from_clipboard(&self) -> Result<Option<String>, PlatformError> {
        let _guard = ClipboardGuard::open()?;
        if unsafe { IsClipboardFormatAvailable(CF_UNICODETEXT.0 as u32) }.is_err() {
            return Ok(None);
        }

        let handle = unsafe { GetClipboardData(CF_UNICODETEXT.0 as u32) }.map_err(|source| {
            PlatformError::Win32 {
                operation: "GetClipboardData",
                source,
            }
        })?;

        let ptr =
            unsafe { GlobalLock(windows::Win32::Foundation::HGLOBAL(handle.0)) } as *const u16;
        if ptr.is_null() {
            return Err(Self::win32_error("GlobalLock"));
        }

        let size_bytes = unsafe { GlobalSize(windows::Win32::Foundation::HGLOBAL(handle.0)) };
        let max_units = size_bytes / size_of::<u16>();
        let slice = unsafe { std::slice::from_raw_parts(ptr, max_units) };
        let len = slice
            .iter()
            .position(|value| *value == 0)
            .unwrap_or(max_units);
        let text =
            String::from_utf16(&slice[..len]).map_err(|_| PlatformError::InvalidClipboardText);
        unsafe {
            let _ = GlobalUnlock(windows::Win32::Foundation::HGLOBAL(handle.0));
        }

        text.map(Some)
    }

    fn write_text_to_clipboard(&self, text: &str) -> Result<(), PlatformError> {
        let _guard = ClipboardGuard::open()?;
        unsafe { EmptyClipboard() }.map_err(|source| PlatformError::Win32 {
            operation: "EmptyClipboard",
            source,
        })?;

        let mut wide = text.encode_utf16().collect::<Vec<_>>();
        wide.push(0);
        let bytes = wide.len() * size_of::<u16>();

        let handle = unsafe { GlobalAlloc(GMEM_MOVEABLE, bytes) }.map_err(|source| {
            PlatformError::Win32 {
                operation: "GlobalAlloc",
                source,
            }
        })?;
        if handle.is_invalid() {
            return Err(Self::win32_error("GlobalAlloc"));
        }

        let ptr = unsafe { GlobalLock(handle) } as *mut u8;
        if ptr.is_null() {
            return Err(Self::win32_error("GlobalLock"));
        }

        unsafe {
            std::ptr::copy_nonoverlapping(wide.as_ptr() as *const u8, ptr, bytes);
            let _ = GlobalUnlock(handle);
        }

        let set = unsafe { SetClipboardData(CF_UNICODETEXT.0 as u32, Some(HANDLE(handle.0))) }
            .map_err(|source| PlatformError::Win32 {
                operation: "SetClipboardData",
                source,
            })?;
        if set.is_invalid() {
            return Err(Self::win32_error("SetClipboardData"));
        }

        Ok(())
    }

    fn read_image_dib_from_clipboard(&self) -> Result<Option<Vec<u8>>, PlatformError> {
        let _guard = ClipboardGuard::open()?;
        if unsafe { IsClipboardFormatAvailable(CF_DIB.0 as u32) }.is_err() {
            if unsafe { IsClipboardFormatAvailable(CF_BITMAP.0 as u32) }.is_err() {
                return Ok(None);
            }

            return read_bitmap_clipboard_as_dib();
        }

        let handle = unsafe { GetClipboardData(CF_DIB.0 as u32) }.map_err(|source| {
            PlatformError::Win32 {
                operation: "GetClipboardData(CF_DIB)",
                source,
            }
        })?;

        let global = windows::Win32::Foundation::HGLOBAL(handle.0);
        let ptr = unsafe { GlobalLock(global) } as *const u8;
        if ptr.is_null() {
            return Err(Self::win32_error("GlobalLock(CF_DIB)"));
        }

        let size_bytes = unsafe { GlobalSize(global) };
        let bytes = unsafe { std::slice::from_raw_parts(ptr, size_bytes) }.to_vec();
        unsafe {
            let _ = GlobalUnlock(global);
        }

        if bytes.is_empty() {
            Ok(None)
        } else {
            Ok(Some(bytes))
        }
    }

    fn read_file_paths_from_clipboard(&self) -> Result<Option<Vec<String>>, PlatformError> {
        let _guard = ClipboardGuard::open()?;
        if unsafe { IsClipboardFormatAvailable(CF_HDROP.0 as u32) }.is_err() {
            return Ok(None);
        }

        let handle = unsafe { GetClipboardData(CF_HDROP.0 as u32) }.map_err(|source| {
            PlatformError::Win32 {
                operation: "GetClipboardData(CF_HDROP)",
                source,
            }
        })?;
        if handle.is_invalid() {
            return Ok(None);
        }

        let hdrop = HDROP(handle.0);
        let count = unsafe { DragQueryFileW(hdrop, u32::MAX, None) };
        if count == 0 {
            return Ok(None);
        }

        let mut paths = Vec::with_capacity(count as usize);
        for index in 0..count {
            let len = unsafe { DragQueryFileW(hdrop, index, None) };
            if len == 0 {
                continue;
            }

            let mut buffer = vec![0u16; len as usize + 1];
            let copied = unsafe { DragQueryFileW(hdrop, index, Some(&mut buffer)) };
            if copied == 0 {
                continue;
            }

            paths.push(String::from_utf16_lossy(&buffer[..copied as usize]));
        }

        if paths.is_empty() {
            Ok(None)
        } else {
            Ok(Some(paths))
        }
    }

    fn write_image_dib_to_clipboard(&self, dib_bytes: &[u8]) -> Result<(), PlatformError> {
        if dib_bytes.is_empty() {
            return Ok(());
        }

        let _guard = ClipboardGuard::open()?;
        unsafe { EmptyClipboard() }.map_err(|source| PlatformError::Win32 {
            operation: "EmptyClipboard",
            source,
        })?;

        let handle = unsafe { GlobalAlloc(GMEM_MOVEABLE, dib_bytes.len()) }.map_err(|source| {
            PlatformError::Win32 {
                operation: "GlobalAlloc(CF_DIB)",
                source,
            }
        })?;
        if handle.is_invalid() {
            return Err(Self::win32_error("GlobalAlloc(CF_DIB)"));
        }

        let ptr = unsafe { GlobalLock(handle) } as *mut u8;
        if ptr.is_null() {
            return Err(Self::win32_error("GlobalLock(CF_DIB)"));
        }

        unsafe {
            std::ptr::copy_nonoverlapping(dib_bytes.as_ptr(), ptr, dib_bytes.len());
            let _ = GlobalUnlock(handle);
        }

        let set = unsafe { SetClipboardData(CF_DIB.0 as u32, Some(HANDLE(handle.0))) }.map_err(
            |source| PlatformError::Win32 {
                operation: "SetClipboardData(CF_DIB)",
                source,
            },
        )?;
        if set.is_invalid() {
            return Err(Self::win32_error("SetClipboardData(CF_DIB)"));
        }

        Ok(())
    }

    fn write_file_paths_to_clipboard(&self, file_paths: &[String]) -> Result<(), PlatformError> {
        let mut wide_paths = Vec::new();
        for path in file_paths
            .iter()
            .map(|path| path.trim())
            .filter(|path| !path.is_empty())
        {
            wide_paths.extend(path.encode_utf16());
            wide_paths.push(0);
        }
        if wide_paths.is_empty() {
            return Ok(());
        }
        wide_paths.push(0);

        let dropfiles_size = size_of::<DROPFILES>();
        let bytes = dropfiles_size + wide_paths.len() * size_of::<u16>();

        let _guard = ClipboardGuard::open()?;
        unsafe { EmptyClipboard() }.map_err(|source| PlatformError::Win32 {
            operation: "EmptyClipboard",
            source,
        })?;

        let handle =
            unsafe { GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, bytes) }.map_err(|source| {
                PlatformError::Win32 {
                    operation: "GlobalAlloc(CF_HDROP)",
                    source,
                }
            })?;
        if handle.is_invalid() {
            return Err(Self::win32_error("GlobalAlloc(CF_HDROP)"));
        }

        let ptr = unsafe { GlobalLock(handle) } as *mut u8;
        if ptr.is_null() {
            return Err(Self::win32_error("GlobalLock(CF_HDROP)"));
        }

        unsafe {
            std::ptr::write(
                ptr as *mut DROPFILES,
                DROPFILES {
                    pFiles: dropfiles_size as u32,
                    pt: POINT::default(),
                    fNC: BOOL(0),
                    fWide: BOOL(1),
                },
            );
            std::ptr::copy_nonoverlapping(
                wide_paths.as_ptr() as *const u8,
                ptr.add(dropfiles_size),
                wide_paths.len() * size_of::<u16>(),
            );
            let _ = GlobalUnlock(handle);
        }

        let set = unsafe { SetClipboardData(CF_HDROP.0 as u32, Some(HANDLE(handle.0))) }.map_err(
            |source| PlatformError::Win32 {
                operation: "SetClipboardData(CF_HDROP)",
                source,
            },
        )?;
        if set.is_invalid() {
            return Err(Self::win32_error("SetClipboardData(CF_HDROP)"));
        }

        Ok(())
    }

    fn paste_into_window(&self, target: WindowHandle) -> Result<(), PlatformError> {
        let target = self
            .resolve_top_level_window(target)
            .ok_or(PlatformError::InvalidWindow)?;
        activate_window(target)?;

        if post_paste_message(target)? {
            return Ok(());
        }

        if send_ctrl_v()? {
            return Ok(());
        }

        if send_shift_insert()? {
            return Ok(());
        }

        Err(Self::win32_error("paste injection"))
    }

    fn client_area_animations_enabled(&self) -> bool {
        let mut enabled = BOOL(1);
        unsafe {
            SystemParametersInfoW(
                SPI_GETCLIENTAREAANIMATION,
                0,
                Some((&mut enabled as *mut BOOL).cast::<core::ffi::c_void>()),
                SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS::default(),
            )
        }
        .is_ok()
            && enabled.as_bool()
    }

    fn find_window_by_title(&self, title: &str) -> Result<Option<WindowHandle>, PlatformError> {
        Ok(find_current_process_window_by_title(title))
    }

    fn show_window(&self, window: WindowHandle) -> Result<(), PlatformError> {
        ensure_window(window)?;
        unsafe {
            let _ = ShowWindow(window.hwnd(), SW_RESTORE);
        }
        activate_window(window)
    }

    fn hide_window(&self, window: WindowHandle) -> Result<(), PlatformError> {
        ensure_window(window)?;
        unsafe {
            let _ = ShowWindow(window.hwnd(), SW_HIDE);
        }
        Ok(())
    }

    fn move_window(
        &self,
        window: WindowHandle,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
    ) -> Result<(), PlatformError> {
        ensure_window(window)?;
        unsafe {
            SetWindowPos(
                window.hwnd(),
                Some(HWND_TOPMOST),
                x,
                y,
                width,
                height,
                SWP_NOACTIVATE | SWP_NOOWNERZORDER,
            )
        }
        .map_err(|source| PlatformError::Win32 {
            operation: "SetWindowPos",
            source,
        })
    }

    fn add_tray_icon(&self, window: WindowHandle) -> Result<(), PlatformError> {
        ensure_window(window)?;
        let icon = unsafe { LoadIconW(None, IDI_APPLICATION) }.unwrap_or(HICON::default());
        let mut data = notify_icon_data(window, icon);
        let added = unsafe { Shell_NotifyIconW(NIM_ADD, &data).as_bool() };
        if !added {
            return Err(Self::win32_error("Shell_NotifyIconW(NIM_ADD)"));
        }

        data.Anonymous = NOTIFYICONDATAW_0 {
            uVersion: NOTIFYICON_VERSION_4,
        };
        let _ = unsafe { Shell_NotifyIconW(NIM_SETVERSION, &data) };

        Ok(())
    }

    fn remove_tray_icon(&self, window: WindowHandle) -> Result<(), PlatformError> {
        if window.is_null() {
            return Ok(());
        }

        let data = notify_icon_data(window, HICON::default());
        if unsafe { Shell_NotifyIconW(NIM_DELETE, &data).as_bool() } {
            Ok(())
        } else {
            Err(Self::win32_error("Shell_NotifyIconW(NIM_DELETE)"))
        }
    }
}

#[cfg(windows)]
static PLATFORM_EVENT_SENDER: OnceLock<Mutex<Option<mpsc::Sender<PlatformEvent>>>> =
    OnceLock::new();

#[cfg(windows)]
fn spawn_windows_event_loop(
) -> Result<(PlatformEventLoop, mpsc::Receiver<PlatformEvent>), PlatformError> {
    let (event_tx, event_rx) = mpsc::channel();
    let (ready_tx, ready_rx) = mpsc::channel();
    PLATFORM_EVENT_SENDER
        .get_or_init(|| Mutex::new(None))
        .lock()
        .expect("platform event sender mutex poisoned")
        .replace(event_tx);

    let thread = thread::Builder::new()
        .name("paste-win32-message-loop".to_string())
        .spawn(move || {
            PLATFORM_FEATURE_STATUS.store(0, Ordering::SeqCst);
            PLATFORM_TRAY_LAST_ERROR.store(0, Ordering::SeqCst);
            let startup_result = unsafe { create_message_window() };
            let Ok(hwnd) = startup_result else {
                let _ = ready_tx.send(Err(startup_result.err().unwrap()));
                clear_event_sender();
                return;
            };

            let window = WindowHandle::from_hwnd(hwnd);
            let platform = WindowsPlatformIntegration::new();
            if platform.register_overlay_hotkey(window).is_ok() {
                PLATFORM_FEATURE_STATUS.fetch_or(PLATFORM_STATUS_HOTKEY, Ordering::SeqCst);
            }
            if platform.add_clipboard_listener(window).is_ok() {
                PLATFORM_FEATURE_STATUS
                    .fetch_or(PLATFORM_STATUS_CLIPBOARD_LISTENER, Ordering::SeqCst);
            }
            match platform.add_tray_icon(window) {
                Ok(()) => {
                    PLATFORM_FEATURE_STATUS.fetch_or(PLATFORM_STATUS_TRAY_ICON, Ordering::SeqCst);
                }
                Err(error) => {
                    PLATFORM_TRAY_LAST_ERROR.store(platform_error_code(&error), Ordering::SeqCst);
                }
            }
            let _ = ready_tx.send(Ok(window));

            unsafe {
                let mut msg = MSG::default();
                while GetMessageW(&mut msg, None, 0, 0).as_bool() {
                    let _ = TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                }
            }

            let _ = platform.remove_clipboard_listener(window);
            let _ = platform.unregister_overlay_hotkey(window);
            let _ = platform.remove_tray_icon(window);
            send_platform_event(PlatformEvent::MessageLoopStopped);
            clear_event_sender();
        })
        .map_err(|_| PlatformError::EventLoopStartup)?;

    let window = ready_rx
        .recv()
        .map_err(|_| PlatformError::EventLoopStartup)??;
    Ok((
        PlatformEventLoop {
            window,
            thread: Some(thread),
        },
        event_rx,
    ))
}

#[cfg(windows)]
unsafe fn create_message_window() -> Result<HWND, PlatformError> {
    let class_name = w!("PasteGPUIMessageWindow");
    let module = GetModuleHandleW(PCWSTR::null()).map_err(|source| PlatformError::Win32 {
        operation: "GetModuleHandleW",
        source,
    })?;
    let hinstance = HINSTANCE(module.0);

    let window_class = WNDCLASSW {
        lpfnWndProc: Some(message_window_proc),
        hInstance: hinstance,
        lpszClassName: class_name,
        ..Default::default()
    };

    let atom = RegisterClassW(&window_class);
    if atom == 0 {
        return Err(WindowsPlatformIntegration::win32_error("RegisterClassW"));
    }

    CreateWindowExW(
        WINDOW_EX_STYLE::default(),
        class_name,
        w!("Paste GPUI Message Window"),
        WINDOW_STYLE::default(),
        0,
        0,
        0,
        0,
        None,
        None,
        Some(hinstance),
        None,
    )
    .map_err(|source| PlatformError::Win32 {
        operation: "CreateWindowExW",
        source,
    })
}

#[cfg(windows)]
unsafe extern "system" fn message_window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_HOTKEY if wparam.0 as i32 == OVERLAY_HOTKEY_ID => {
            send_platform_event(PlatformEvent::OverlayHotkeyPressed);
            LRESULT(0)
        }
        WM_CLIPBOARDUPDATE => {
            send_platform_event(PlatformEvent::ClipboardUpdated);
            LRESULT(0)
        }
        WM_TRAY_ICON => {
            let notification = (lparam.0 as u32) & 0xffff;
            match notification {
                NIN_SELECT | NIN_KEYSELECT | WM_LBUTTONUP => {
                    send_platform_event(PlatformEvent::TrayShowRequested);
                }
                WM_CONTEXTMENU => {
                    show_tray_menu(hwnd);
                }
                _ => {}
            }
            LRESULT(0)
        }
        WM_PASTE_GPUI_TEST_TRAY_COMMAND => {
            handle_tray_command(hwnd, wparam.0);
            LRESULT(0)
        }
        WM_PASTE_GPUI_TEST_PLATFORM_STATUS => {
            LRESULT(PLATFORM_FEATURE_STATUS.load(Ordering::SeqCst) as isize)
        }
        WM_PASTE_GPUI_TEST_TRAY_ERROR => {
            LRESULT(PLATFORM_TRAY_LAST_ERROR.load(Ordering::SeqCst) as isize)
        }
        WM_CLOSE => {
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            PostQuitMessage(0);
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

#[cfg(windows)]
fn send_platform_event(event: PlatformEvent) {
    if let Some(sender) = PLATFORM_EVENT_SENDER
        .get()
        .and_then(|sender| sender.lock().ok())
        .and_then(|sender| sender.as_ref().cloned())
    {
        let _ = sender.send(event);
    }
}

#[cfg(windows)]
fn clear_event_sender() {
    if let Some(sender) = PLATFORM_EVENT_SENDER.get() {
        if let Ok(mut sender) = sender.lock() {
            sender.take();
        }
    }
}

#[cfg(windows)]
fn platform_error_code(error: &PlatformError) -> usize {
    match error {
        PlatformError::Win32 { source, .. } => source.code().0 as u32 as usize,
        PlatformError::InvalidWindow => 0xFFFF_FF01,
        PlatformError::InvalidClipboardText => 0xFFFF_FF02,
        PlatformError::EventLoopStartup => 0xFFFF_FF03,
        PlatformError::NotImplemented => 0xFFFF_FF04,
    }
}

#[cfg(windows)]
unsafe fn show_tray_menu(hwnd: HWND) {
    let Ok(menu) = CreatePopupMenu() else {
        return;
    };

    let _ = AppendMenuW(menu, MF_STRING, TRAY_MENU_SHOW, w!("Show"));
    let _ = AppendMenuW(menu, MF_STRING, TRAY_MENU_RESTART, w!("Restart"));
    let _ = AppendMenuW(menu, MF_STRING, TRAY_MENU_EXIT, w!("Exit"));

    let mut point = POINT::default();
    if GetCursorPos(&mut point).is_err() {
        let _ = DestroyMenu(menu);
        return;
    }

    let _ = SetForegroundWindow(hwnd);
    let command = TrackPopupMenuEx(
        menu,
        TPM_LEFTALIGN.0 | TPM_BOTTOMALIGN.0 | TPM_RIGHTBUTTON.0 | TPM_RETURNCMD.0,
        point.x,
        point.y,
        hwnd,
        None,
    );
    let _ = DestroyMenu(menu);

    handle_tray_command(hwnd, command.0 as usize);
}

#[cfg(windows)]
fn handle_tray_command(hwnd: HWND, command: usize) {
    match command {
        TRAY_MENU_SHOW | TEST_TRAY_COMMAND_SHOW => {
            send_platform_event(PlatformEvent::TrayShowRequested)
        }
        TRAY_MENU_RESTART | TEST_TRAY_COMMAND_RESTART => {
            send_platform_event(PlatformEvent::TrayRestartRequested)
        }
        TRAY_MENU_EXIT | TEST_TRAY_COMMAND_EXIT => {
            send_platform_event(PlatformEvent::TrayExitRequested);
            unsafe {
                let _ = PostMessageW(Some(hwnd), WM_CLOSE, WPARAM(0), LPARAM(0));
            }
        }
        _ => {}
    }
}

#[cfg(windows)]
fn notify_icon_data(window: WindowHandle, icon: HICON) -> NOTIFYICONDATAW {
    let mut data = NOTIFYICONDATAW {
        cbSize: size_of::<NOTIFYICONDATAW>() as u32,
        hWnd: window.hwnd(),
        uID: TRAY_ICON_ID,
        uFlags: NOTIFY_ICON_DATA_FLAGS(NIF_MESSAGE.0 | NIF_ICON.0 | NIF_TIP.0),
        uCallbackMessage: WM_TRAY_ICON,
        hIcon: icon,
        ..Default::default()
    };

    let tip = "Paste GPUI";
    for (index, code_unit) in tip.encode_utf16().take(data.szTip.len() - 1).enumerate() {
        data.szTip[index] = code_unit;
    }

    data
}

#[cfg(windows)]
struct WindowSearch<'a> {
    process_id: u32,
    title: &'a str,
    found: Option<HWND>,
}

#[cfg(windows)]
fn find_current_process_window_by_title(title: &str) -> Option<WindowHandle> {
    let mut search = WindowSearch {
        process_id: unsafe { GetCurrentProcessId() },
        title,
        found: None,
    };

    unsafe {
        let _ = EnumWindows(
            Some(enum_window_by_title),
            LPARAM((&mut search as *mut WindowSearch<'_>) as isize),
        );
    }

    search.found.map(WindowHandle::from_hwnd)
}

#[cfg(windows)]
unsafe extern "system" fn enum_window_by_title(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let search = &mut *(lparam.0 as *mut WindowSearch<'_>);
    let mut process_id = 0;
    GetWindowThreadProcessId(hwnd, Some(&mut process_id));
    if process_id != search.process_id {
        return BOOL(1);
    }

    if window_title(hwnd).as_deref() == Some(search.title) {
        search.found = Some(hwnd);
        return BOOL(0);
    }

    BOOL(1)
}

#[cfg(windows)]
fn window_title(hwnd: HWND) -> Option<String> {
    let len = unsafe { GetWindowTextLengthW(hwnd) };
    if len <= 0 {
        return None;
    }

    let mut buffer = vec![0u16; len as usize + 1];
    let copied = unsafe { GetWindowTextW(hwnd, &mut buffer) };
    if copied <= 0 {
        return None;
    }

    Some(String::from_utf16_lossy(&buffer[..copied as usize]))
}

#[cfg(windows)]
fn read_bitmap_clipboard_as_dib() -> Result<Option<Vec<u8>>, PlatformError> {
    let handle =
        unsafe { GetClipboardData(CF_BITMAP.0 as u32) }.map_err(|source| PlatformError::Win32 {
            operation: "GetClipboardData(CF_BITMAP)",
            source,
        })?;
    if handle.is_invalid() {
        return Ok(None);
    }

    let bitmap_handle = HBITMAP(handle.0);
    let mut bitmap = BITMAP::default();
    let object_size = unsafe {
        GetObjectW(
            HGDIOBJ(bitmap_handle.0),
            size_of::<BITMAP>() as i32,
            Some((&mut bitmap as *mut BITMAP).cast()),
        )
    };
    if object_size == 0 || bitmap.bmWidth <= 0 || bitmap.bmHeight <= 0 {
        return Err(WindowsPlatformIntegration::win32_error(
            "GetObjectW(CF_BITMAP)",
        ));
    }

    let width = bitmap.bmWidth;
    let height = bitmap.bmHeight;
    let bit_count = 24u16;
    let stride = ((width as usize * bit_count as usize + 31) / 32) * 4;
    let pixel_bytes = stride * height as usize;

    let mut info = BITMAPINFO {
        bmiHeader: BITMAPINFOHEADER {
            biSize: size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: width,
            biHeight: height,
            biPlanes: 1,
            biBitCount: bit_count,
            biCompression: BI_RGB.0,
            biSizeImage: pixel_bytes as u32,
            ..Default::default()
        },
        ..Default::default()
    };
    let mut pixels = vec![0u8; pixel_bytes];

    let dc = unsafe { GetDC(None) };
    if dc.is_invalid() {
        return Err(WindowsPlatformIntegration::win32_error("GetDC"));
    }

    let scan_lines = unsafe {
        GetDIBits(
            dc,
            bitmap_handle,
            0,
            height as u32,
            Some(pixels.as_mut_ptr().cast()),
            &mut info,
            DIB_RGB_COLORS,
        )
    };
    unsafe {
        let _ = ReleaseDC(None, dc);
    }
    if scan_lines == 0 {
        return Err(WindowsPlatformIntegration::win32_error(
            "GetDIBits(CF_BITMAP)",
        ));
    }

    let mut dib = Vec::with_capacity(size_of::<BITMAPINFOHEADER>() + pixels.len());
    dib.extend_from_slice(&info.bmiHeader.biSize.to_le_bytes());
    dib.extend_from_slice(&info.bmiHeader.biWidth.to_le_bytes());
    dib.extend_from_slice(&info.bmiHeader.biHeight.to_le_bytes());
    dib.extend_from_slice(&info.bmiHeader.biPlanes.to_le_bytes());
    dib.extend_from_slice(&info.bmiHeader.biBitCount.to_le_bytes());
    dib.extend_from_slice(&info.bmiHeader.biCompression.to_le_bytes());
    dib.extend_from_slice(&info.bmiHeader.biSizeImage.to_le_bytes());
    dib.extend_from_slice(&info.bmiHeader.biXPelsPerMeter.to_le_bytes());
    dib.extend_from_slice(&info.bmiHeader.biYPelsPerMeter.to_le_bytes());
    dib.extend_from_slice(&info.bmiHeader.biClrUsed.to_le_bytes());
    dib.extend_from_slice(&info.bmiHeader.biClrImportant.to_le_bytes());
    dib.extend_from_slice(&pixels);

    Ok(Some(dib))
}

#[cfg(not(windows))]
impl PlatformIntegration for WindowsPlatformIntegration {
    fn register_overlay_hotkey(&self, _window: WindowHandle) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn unregister_overlay_hotkey(&self, _window: WindowHandle) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn add_clipboard_listener(&self, _window: WindowHandle) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn remove_clipboard_listener(&self, _window: WindowHandle) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn foreground_window(&self) -> Option<WindowHandle> {
        None
    }

    fn resolve_top_level_window(&self, _window: WindowHandle) -> Option<WindowHandle> {
        None
    }

    fn window_process_name(&self, _window: WindowHandle) -> Result<Option<String>, PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn is_window(&self, _window: WindowHandle) -> bool {
        false
    }

    fn read_text_from_clipboard(&self) -> Result<Option<String>, PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn read_image_dib_from_clipboard(&self) -> Result<Option<Vec<u8>>, PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn read_file_paths_from_clipboard(&self) -> Result<Option<Vec<String>>, PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn write_text_to_clipboard(&self, _text: &str) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn write_image_dib_to_clipboard(&self, _dib_bytes: &[u8]) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn write_file_paths_to_clipboard(&self, _file_paths: &[String]) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn paste_into_window(&self, _target: WindowHandle) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn client_area_animations_enabled(&self) -> bool {
        true
    }

    fn find_window_by_title(&self, _title: &str) -> Result<Option<WindowHandle>, PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn show_window(&self, _window: WindowHandle) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn hide_window(&self, _window: WindowHandle) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn move_window(
        &self,
        _window: WindowHandle,
        _x: i32,
        _y: i32,
        _width: i32,
        _height: i32,
    ) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn add_tray_icon(&self, _window: WindowHandle) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }

    fn remove_tray_icon(&self, _window: WindowHandle) -> Result<(), PlatformError> {
        Err(PlatformError::NotImplemented)
    }
}

#[cfg(windows)]
struct ClipboardGuard;

#[cfg(windows)]
impl ClipboardGuard {
    fn open() -> Result<Self, PlatformError> {
        unsafe { OpenClipboard(None) }.map_err(|source| PlatformError::Win32 {
            operation: "OpenClipboard",
            source,
        })?;
        Ok(Self)
    }
}

#[cfg(windows)]
impl Drop for ClipboardGuard {
    fn drop(&mut self) {
        unsafe {
            let _ = CloseClipboard();
        }
    }
}

#[cfg(windows)]
fn ensure_window(window: WindowHandle) -> Result<(), PlatformError> {
    if window.is_null() || !unsafe { IsWindow(Some(window.hwnd())).as_bool() } {
        Err(PlatformError::InvalidWindow)
    } else {
        Ok(())
    }
}

#[cfg(windows)]
fn activate_window(target: WindowHandle) -> Result<(), PlatformError> {
    ensure_window(target)?;
    if unsafe { IsIconic(target.hwnd()).as_bool() } {
        unsafe {
            let _ = ShowWindow(target.hwnd(), SW_RESTORE);
        }
    }

    unsafe {
        let _ = AllowSetForegroundWindow(u32::MAX);
    }

    let foreground = unsafe { GetForegroundWindow() };
    let current_thread = unsafe { GetCurrentThreadId() };
    let target_thread = unsafe { GetWindowThreadProcessId(target.hwnd(), None) };
    let foreground_thread = if foreground.0.is_null() {
        0
    } else {
        unsafe { GetWindowThreadProcessId(foreground, None) }
    };

    let attached_target = target_thread != 0
        && target_thread != current_thread
        && unsafe { AttachThreadInput(current_thread, target_thread, true).as_bool() };
    let attached_foreground = foreground_thread != 0
        && foreground_thread != current_thread
        && foreground_thread != target_thread
        && unsafe { AttachThreadInput(current_thread, foreground_thread, true).as_bool() };

    unsafe {
        let _ = BringWindowToTop(target.hwnd());
        let _ = SetForegroundWindow(target.hwnd());
        let _ = SetActiveWindow(target.hwnd());
        focus_target_thread_control(target_thread);
    }

    if attached_foreground {
        unsafe {
            let _ = AttachThreadInput(current_thread, foreground_thread, false);
        }
    }
    if attached_target {
        unsafe {
            let _ = AttachThreadInput(current_thread, target_thread, false);
        }
    }

    Ok(())
}

#[cfg(windows)]
unsafe fn focus_target_thread_control(target_thread: u32) {
    if target_thread == 0 {
        return;
    }

    let mut info = GUITHREADINFO {
        cbSize: size_of::<GUITHREADINFO>() as u32,
        ..Default::default()
    };
    if GetGUIThreadInfo(target_thread, &mut info).is_ok()
        && IsWindow(Some(info.hwndFocus)).as_bool()
    {
        let _ = SetFocus(Some(info.hwndFocus));
    }
}

#[cfg(windows)]
fn post_paste_message(target: WindowHandle) -> Result<bool, PlatformError> {
    let destination = focused_control_or_window(target);
    unsafe { PostMessageW(Some(destination), WM_PASTE, WPARAM(0), LPARAM(0)) }
        .map(|_| true)
        .map_err(|source| PlatformError::Win32 {
            operation: "PostMessageW",
            source,
        })
}

#[cfg(windows)]
fn focused_control_or_window(target: WindowHandle) -> HWND {
    let target_thread = unsafe { GetWindowThreadProcessId(target.hwnd(), None) };
    let mut info = GUITHREADINFO {
        cbSize: size_of::<GUITHREADINFO>() as u32,
        ..Default::default()
    };

    if target_thread != 0
        && unsafe { GetGUIThreadInfo(target_thread, &mut info).is_ok() }
        && unsafe { IsWindow(Some(info.hwndFocus)).as_bool() }
    {
        info.hwndFocus
    } else {
        target.hwnd()
    }
}

#[cfg(windows)]
fn send_ctrl_v() -> Result<bool, PlatformError> {
    send_key_chord(&[VK_CONTROL, VK_V])
}

#[cfg(windows)]
fn send_shift_insert() -> Result<bool, PlatformError> {
    send_key_chord(&[VK_SHIFT, VK_INSERT])
}

#[cfg(windows)]
fn send_key_chord(keys: &[VIRTUAL_KEY]) -> Result<bool, PlatformError> {
    let mut inputs = Vec::with_capacity(keys.len() * 2);
    for key in keys {
        inputs.push(key_input(*key, KEYBD_EVENT_FLAGS(0)));
    }
    for key in keys.iter().rev() {
        inputs.push(key_input(*key, KEYEVENTF_KEYUP));
    }

    let sent = unsafe { SendInput(&inputs, size_of::<INPUT>() as i32) };
    Ok(sent == inputs.len() as u32)
}

#[cfg(windows)]
fn key_input(key: VIRTUAL_KEY, flags: KEYBD_EVENT_FLAGS) -> INPUT {
    INPUT {
        r#type: INPUT_KEYBOARD,
        Anonymous: INPUT_0 {
            ki: KEYBDINPUT {
                wVk: key,
                wScan: 0,
                dwFlags: flags,
                time: 0,
                dwExtraInfo: 0,
            },
        },
    }
}

fn process_name_from_path(path: &str) -> Option<String> {
    let file_name = path.rsplit(['\\', '/']).next().unwrap_or(path).trim();
    if file_name.is_empty() {
        return None;
    }

    Some(
        file_name
            .strip_suffix(".exe")
            .or_else(|| file_name.strip_suffix(".EXE"))
            .unwrap_or(file_name)
            .to_string(),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn null_window_handle_is_null() {
        assert!(WindowHandle(0).is_null());
        assert!(!WindowHandle(42).is_null());
    }

    #[test]
    fn hotkey_id_is_stable() {
        assert_eq!(OVERLAY_HOTKEY_ID, 0x5056);
    }

    #[test]
    fn tray_icon_id_is_stable() {
        assert_eq!(TRAY_ICON_ID, 0x5057);
        assert_eq!(WM_TRAY_ICON, 0x8000 + 0x5057);
    }

    #[test]
    fn process_name_from_path_prefers_executable_stem() {
        assert_eq!(
            process_name_from_path(r"C:\Program Files\App\Editor.exe").as_deref(),
            Some("Editor")
        );
        assert_eq!(
            process_name_from_path("/usr/bin/zed").as_deref(),
            Some("zed")
        );
    }
}
