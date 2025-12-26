use std::sync::{Mutex, Once, OnceLock};
use windows::Win32::Foundation::HMODULE;
use windows::Win32::Foundation::{COLORREF, HWND};
use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};
use windows::core::PCWSTR;

// Dark mode colors
pub const DARK_MENU_BG: COLORREF = COLORREF(0x00202020);
pub const DARK_MENU_HOVER: COLORREF = COLORREF(0x00404040);
pub const DARK_MENU_TEXT: COLORREF = COLORREF(0x00FFFFFF);
pub const DARK_MENU_TEXT_DISABLED: COLORREF = COLORREF(0x00808080);
pub const DARK_MENU_BORDER: COLORREF = COLORREF(0x00262624);
pub const DARK_EDITOR_BG: COLORREF = COLORREF(0x001E1E1E);
pub const DARK_EDITOR_TEXT: COLORREF = COLORREF(0x00E0E0E0);
pub const DARK_STATUSBAR_BG: COLORREF = COLORREF(0x00202020);
pub const DARK_STATUSBAR_TEXT: COLORREF = COLORREF(0x00E0E0E0);
pub const DARK_SEPARATOR: COLORREF = COLORREF(0x00404040);

// Light mode colors
#[allow(dead_code)]
pub const LIGHT_EDITOR_BG: COLORREF = COLORREF(0x00FFFFFF);
pub const LIGHT_EDITOR_TEXT: COLORREF = COLORREF(0x00000000);
pub const LIGHT_STATUSBAR_BG: COLORREF = COLORREF(0x00F0F0F0);
pub const LIGHT_STATUSBAR_TEXT: COLORREF = COLORREF(0x00000000);
pub const LIGHT_SEPARATOR: COLORREF = COLORREF(0x00D0D0D0);
pub const LIGHT_MENU_BG: COLORREF = COLORREF(0x00F0F0F0);

// Dark mode state
pub static DARK_MODE_ENABLED: Mutex<bool> = Mutex::new(false);
pub static DARK_MODE_INIT: Once = Once::new();

// UxTheme DLL handle (loaded once and reused)
// Stored as isize to allow static sharing (HMODULE is not Send/Sync)
static UXTHEME_HANDLE: OnceLock<Option<isize>> = OnceLock::new();

// Get UxTheme DLL handle (loads once on first call)
fn get_uxtheme_handle() -> Option<HMODULE> {
    let handle_opt = *UXTHEME_HANDLE.get_or_init(|| unsafe {
        let uxtheme_lib = "uxtheme.dll\0".encode_utf16().collect::<Vec<_>>();
        LoadLibraryW(PCWSTR(uxtheme_lib.as_ptr()))
            .ok()
            .map(|h| h.0 as isize)
    });
    handle_opt.map(|h| HMODULE(h as *mut _))
}

// Detect if system is using dark mode using UxTheme API
pub fn is_system_dark_mode() -> bool {
    unsafe {
        if let Some(huxtheme) = get_uxtheme_handle() {
            // Ordinal 132 = ShouldAppsUseDarkMode (undocumented API)
            if let Some(should_use_dark) =
                GetProcAddress(huxtheme, windows::core::PCSTR(132 as *const u8))
            {
                type ShouldAppsUseDarkMode = unsafe extern "system" fn() -> bool;
                let func: ShouldAppsUseDarkMode = std::mem::transmute(should_use_dark);
                return func();
            }
        }
        false // Default to light mode if can't detect
    }
}

// Initialize dark mode from system settings
pub fn init_dark_mode() {
    DARK_MODE_INIT.call_once(|| {
        let system_dark = is_system_dark_mode();
        if let Ok(mut dark_mode) = DARK_MODE_ENABLED.lock() {
            *dark_mode = system_dark;
        }
    });
}

// Check if dark mode is enabled
pub fn should_use_dark_mode() -> bool {
    init_dark_mode();
    if let Ok(enabled) = DARK_MODE_ENABLED.lock() {
        *enabled
    } else {
        false
    }
}

// Set preferred app mode (0 = default, 1 = dark, 2 = light)
pub fn set_preferred_app_mode(mode: i32) {
    unsafe {
        if let Some(huxtheme) = get_uxtheme_handle() {
            // Ordinal 135 = SetPreferredAppMode
            if let Some(func_addr) =
                GetProcAddress(huxtheme, windows::core::PCSTR(135 as *const u8))
            {
                type SetPreferredAppMode = unsafe extern "system" fn(i32) -> i32;
                let func: SetPreferredAppMode = std::mem::transmute(func_addr);
                func(mode);
            }
        }
    }
}

// Allow dark mode for specific window
pub fn allow_dark_mode_for_window(hwnd: HWND, enabled: bool) {
    unsafe {
        if let Some(huxtheme) = get_uxtheme_handle() {
            // Ordinal 133 = AllowDarkModeForWindow
            if let Some(func_addr) =
                GetProcAddress(huxtheme, windows::core::PCSTR(133 as *const u8))
            {
                type AllowDarkModeForWindow = unsafe extern "system" fn(HWND, i32) -> i32;
                let func: AllowDarkModeForWindow = std::mem::transmute(func_addr);
                func(hwnd, if enabled { 1 } else { 0 });
            }
        }
    }
}

// Flush menu themes to apply changes
pub fn flush_menu_themes() {
    unsafe {
        if let Some(huxtheme) = get_uxtheme_handle() {
            // Ordinal 136 = FlushMenuThemes
            if let Some(func_addr) =
                GetProcAddress(huxtheme, windows::core::PCSTR(136 as *const u8))
            {
                type FlushMenuThemes = unsafe extern "system" fn();
                let func: FlushMenuThemes = std::mem::transmute(func_addr);
                func();
            }
        }
    }
}

// Set window theme (for scrollbars, etc.)
pub fn set_window_theme(hwnd: HWND, dark_mode: bool) {
    unsafe {
        if let Some(huxtheme) = get_uxtheme_handle() {
            let func_name = b"SetWindowTheme\0";
            if let Some(func_addr) =
                GetProcAddress(huxtheme, windows::core::PCSTR(func_name.as_ptr()))
            {
                type SetWindowThemeFn =
                    unsafe extern "system" fn(HWND, *const u16, *const u16) -> i32;
                let func: SetWindowThemeFn = std::mem::transmute(func_addr);

                if dark_mode {
                    let theme = "DarkMode_Explorer\0".encode_utf16().collect::<Vec<_>>();
                    let _ = func(hwnd, theme.as_ptr(), std::ptr::null());
                } else {
                    let theme = "Explorer\0".encode_utf16().collect::<Vec<_>>();
                    let _ = func(hwnd, theme.as_ptr(), std::ptr::null());
                }
            }
        }
    }
}
