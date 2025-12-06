#![windows_subsystem = "windows"]

mod constants;
mod context_menu;
mod file_io;
mod i18n;
mod line_column;
mod status_bar;

use constants::{
    EC_TOPMARGIN, EM_EXLIMITTEXT, EM_GETLANGOPTIONS, EM_GETTEXT, EM_SETLANGOPTIONS,
    EM_SETPARAFORMAT, EM_SETTARGETDEVICE, EM_SETTEXT, ES_MULTILINE, ICON_BIG, ICON_SMALL,
    ID_EDIT_COPY, ID_EDIT_CUT, ID_EDIT_PASTE, ID_EDIT_SELECTALL, ID_FILE_EXIT, ID_FILE_NEW,
    ID_FILE_OPEN, ID_FILE_SAVE, ID_FILE_SAVEAS, ID_VIEW_STATUSBAR, ID_VIEW_WORDWRAP, IMF_AUTOFONT,
    IMF_DUALFONT, OLE_PLACEHOLDER, PFM_LINESPACING, PFM_SPACEAFTER, PFM_SPACEBEFORE,
};
use context_menu::show_context_menu;
use i18n::{get_string, init_language};
use status_bar::{EM_GETSEL, EM_LINEFROMCHAR, update_status_bar};
use std::path::PathBuf;
use std::sync::Mutex;

// Global variables for file state
static CURRENT_FILE: Mutex<Option<PathBuf>> = Mutex::new(None);
static LAST_MODIFIED_STATE: Mutex<bool> = Mutex::new(false);
static WORD_WRAP_ENABLED: Mutex<bool> = Mutex::new(true);
static STATUSBAR_VISIBLE: Mutex<bool> = Mutex::new(true);
static MENU_HANDLE: Mutex<Option<isize>> = Mutex::new(None);
use windows_sys::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows_sys::Win32::Graphics::Gdi::{CreateFontW, GetSysColorBrush, InvalidateRect};
use windows_sys::Win32::System::DataExchange::{CloseClipboard, GetClipboardData, OpenClipboard};
use windows_sys::Win32::System::LibraryLoader::GetModuleHandleW;
use windows_sys::Win32::UI::Controls::{EM_GETMODIFY, EM_SETMARGINS, EM_SETMODIFY, EM_SETSEL};
use windows_sys::Win32::UI::Input::KeyboardAndMouse::GetKeyState;
use windows_sys::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CS_HREDRAW, CS_VREDRAW, CW_USEDEFAULT, CheckMenuItem, CreateMenu, CreateWindowExW,
    DefWindowProcW, DestroyWindow, DispatchMessageW, EC_LEFTMARGIN, GetClientRect, GetCursorPos,
    GetMessageW, GetWindowLongPtrW, GetWindowRect, IDC_ARROW, LoadCursorW, LoadIconW, MF_CHECKED,
    MF_POPUP, MF_STRING, MF_UNCHECKED, MSG, PostQuitMessage, RegisterClassW, SWP_NOZORDER,
    SendMessageW, SetCursor, SetMenu, SetWindowLongPtrW, SetWindowPos, SetWindowTextW, ShowWindow,
    TranslateMessage, WM_CLOSE, WM_COMMAND, WM_CONTEXTMENU, WM_COPY, WM_CREATE, WM_CUT, WM_DESTROY,
    WM_KEYDOWN, WM_NOTIFY, WM_PASTE, WM_SETCURSOR, WM_SETFONT, WM_SETICON, WM_SIZE, WNDCLASSW,
    WS_CHILD, WS_HSCROLL, WS_OVERLAPPEDWINDOW, WS_VISIBLE, WS_VSCROLL,
};

// Helper function to check if file is "Untitled" or "無題"
fn is_untitled_file(path: &PathBuf) -> bool {
    if let Some(filename) = path.file_name() {
        if let Some(name_str) = filename.to_str() {
            return name_str == "Untitled" || name_str == "無題";
        }
    }
    false
}

// Helper function to update title based on modified state
fn update_title_if_needed(hwnd: HWND, edit_hwnd: HWND) {
    unsafe {
        let is_modified = SendMessageW(edit_hwnd, EM_GETMODIFY, 0, 0) != 0;

        // Check if state changed
        if let Ok(mut last_state) = LAST_MODIFIED_STATE.lock() {
            if *last_state == is_modified {
                return; // No change, skip update
            }
            *last_state = is_modified;
        }

        // Update title
        let app_name = get_string("WINDOW_TITLE");
        if let Ok(current_file) = CURRENT_FILE.lock() {
            if let Some(path) = current_file.as_ref() {
                if let Some(filename) = path.file_name() {
                    if let Some(filename_str) = filename.to_str() {
                        let title = if is_modified {
                            format!("*{} - {}\0", filename_str, app_name)
                        } else {
                            format!("{} - {}\0", filename_str, app_name)
                        };
                        let title_utf16: Vec<u16> = title.encode_utf16().collect();
                        SetWindowTextW(hwnd, title_utf16.as_ptr());
                    }
                }
            }
        }
    }
}

// Helper function to remove OLE objects from RichEdit
fn remove_ole_objects(edit_hwnd: HWND) {
    unsafe {
        // Get text length
        let text_length = SendMessageW(edit_hwnd, 0x000E, 0, 0) as i32; // WM_GETTEXTLENGTH

        if text_length <= 0 {
            return;
        }

        // Allocate buffer and get text
        let mut buffer = vec![0u16; (text_length + 1) as usize];
        let actual_len = SendMessageW(
            edit_hwnd,
            EM_GETTEXT,
            (text_length + 1) as usize,
            buffer.as_mut_ptr() as isize,
        ) as usize;

        if actual_len == 0 {
            return;
        }

        // Check for OLE placeholder characters (0xFFFC = standard OLE placeholder, 0x0001 = alternative)
        let mut has_ole = false;
        buffer.retain(|&ch| {
            if ch == 0xFFFC || ch == OLE_PLACEHOLDER || ch == 0xFFFD {
                has_ole = true;
                false // Remove OLE placeholder
            } else {
                true
            }
        });

        // If OLE objects were found and removed, update the text
        if has_ole {
            let new_text: Vec<u16> = buffer.iter().copied().chain(std::iter::once(0)).collect();
            SendMessageW(edit_hwnd, EM_SETTEXT, 0, new_text.as_ptr() as isize);
        }
    }
}

// Helper function to toggle word wrap
fn toggle_word_wrap(edit_hwnd: HWND) {
    let new_state = {
        if let Ok(mut wrap_state) = WORD_WRAP_ENABLED.lock() {
            *wrap_state = !*wrap_state;
            *wrap_state
        } else {
            return;
        }
    };

    unsafe {
        if new_state {
            // Enable word wrap: set target device to width of client area
            SendMessageW(edit_hwnd, EM_SETTARGETDEVICE, 0, 0);
        } else {
            // Disable word wrap: set target device to null (no wrapping)
            SendMessageW(edit_hwnd, EM_SETTARGETDEVICE, 0, 1);
        }
    }

    // Update menu check state (lock released before this call)
    update_wordwrap_menu_check();
}

// Helper function to update word wrap menu check state
fn update_wordwrap_menu_check() {
    if let Ok(menu_handle) = MENU_HANDLE.lock() {
        if let Some(hmenu_isize) = *menu_handle {
            if let Ok(wrap_state) = WORD_WRAP_ENABLED.lock() {
                let check_state = if *wrap_state {
                    MF_CHECKED
                } else {
                    MF_UNCHECKED
                };
                unsafe {
                    CheckMenuItem(
                        hmenu_isize as *mut std::ffi::c_void,
                        ID_VIEW_WORDWRAP as u32,
                        check_state,
                    );
                }
            }
        }
    }
}

// Helper function to update status bar menu check state
fn update_statusbar_menu_check() {
    if let Ok(menu_handle) = MENU_HANDLE.lock() {
        if let Some(hmenu_isize) = *menu_handle {
            if let Ok(visible) = STATUSBAR_VISIBLE.lock() {
                let check_state = if *visible { MF_CHECKED } else { MF_UNCHECKED };
                unsafe {
                    CheckMenuItem(
                        hmenu_isize as *mut std::ffi::c_void,
                        ID_VIEW_STATUSBAR as u32,
                        check_state,
                    );
                }
            }
        }
    }
}

extern "system" fn window_proc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let hinstance = GetModuleHandleW(std::ptr::null());

                // Load RichEdit library (MSFTEDIT.DLL for RichEdit 4.1+)
                let richedit_lib = "Msftedit.dll\0".encode_utf16().collect::<Vec<_>>();
                let _ =
                    windows_sys::Win32::System::LibraryLoader::LoadLibraryW(richedit_lib.as_ptr());

                let richedit_class = "RICHEDIT50W\0".encode_utf16().collect::<Vec<_>>();

                let edit_hwnd = CreateWindowExW(
                    0,
                    richedit_class.as_ptr(),
                    std::ptr::null(),
                    WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL | ES_MULTILINE,
                    0,
                    0,
                    0,
                    0,
                    hwnd,
                    std::ptr::null_mut(),
                    hinstance,
                    std::ptr::null_mut(),
                );

                // Store the RichEdit control handle in window extra bytes
                SetWindowLongPtrW(hwnd, 0, edit_hwnd as isize);

                // Set unlimited text size (EM_EXLIMITTEXT with max u64 value)
                SendMessageW(edit_hwnd, EM_EXLIMITTEXT, 0, u64::MAX as isize);

                // Enable word wrap by default
                SendMessageW(edit_hwnd, EM_SETTARGETDEVICE, 0, 0);

                // Set left margin (8 pixels)
                let margin = (8u32) | ((0u32) << 16);
                SendMessageW(
                    edit_hwnd,
                    EM_SETMARGINS,
                    EC_LEFTMARGIN as usize,
                    margin as isize,
                );

                // Set top margin (8 pixels)
                let margin = (8u32) | ((0u32) << 16);
                SendMessageW(
                    edit_hwnd,
                    EM_SETMARGINS,
                    EC_TOPMARGIN as usize,
                    margin as isize,
                );

                // Set default font for RichEdit
                let font_name = "MS Gothic\0".encode_utf16().collect::<Vec<_>>();
                let hfont_edit = CreateFontW(
                    -16, // 12pt
                    0,
                    0,
                    0,
                    400, // FW_NORMAL
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    font_name.as_ptr(),
                );
                SendMessageW(edit_hwnd, WM_SETFONT, hfont_edit as usize, 1);

                // Disable auto font (IMF_AUTOFONT | IMF_DUALFONT) to prevent font changes
                let lang_options = SendMessageW(edit_hwnd, EM_GETLANGOPTIONS, 0, 0);
                let new_options = lang_options & !(IMF_AUTOFONT | IMF_DUALFONT) as isize;
                SendMessageW(edit_hwnd, EM_SETLANGOPTIONS, 0, new_options);

                // Set reduced line spacing using PARAFORMAT2
                #[repr(C)]
                #[allow(non_snake_case)]
                struct PARAFORMAT2 {
                    cbSize: u32,
                    dwMask: u32,
                    wNumbering: u16,
                    wReserved: u16,
                    dxStartIndent: i32,
                    dxRightIndent: i32,
                    dxOffset: i32,
                    wAlignment: u16,
                    cTabCount: i16,
                    rgxTabs: [i32; 32],
                    dySpaceBefore: i32,
                    dySpaceAfter: i32,
                    dyLineSpacing: i32,
                    sStyle: i16,
                    bLineSpacingRule: u8,
                    bOutlineLevel: u8,
                    wShadingWeight: u16,
                    wShadingStyle: u16,
                    wNumberingStart: u16,
                    wNumberingStyle: u16,
                    wNumberingTab: u16,
                    wBorderSpace: u16,
                    wBorderWidth: u16,
                    wBorders: u16,
                }

                let pf = PARAFORMAT2 {
                    cbSize: std::mem::size_of::<PARAFORMAT2>() as u32,
                    dwMask: PFM_LINESPACING | PFM_SPACEBEFORE | PFM_SPACEAFTER,
                    wNumbering: 0,
                    wReserved: 0,
                    dxStartIndent: 0,
                    dxRightIndent: 0,
                    dxOffset: 0,
                    wAlignment: 0,
                    cTabCount: 0,
                    rgxTabs: [0; 32],
                    dySpaceBefore: 0,
                    dySpaceAfter: 0,
                    dyLineSpacing: 260, // 13pt in twips (260 = 13 * 20)
                    sStyle: 0,
                    bLineSpacingRule: 4, // Rule 4: Exact spacing
                    bOutlineLevel: 0,
                    wShadingWeight: 0,
                    wShadingStyle: 0,
                    wNumberingStart: 0,
                    wNumberingStyle: 0,
                    wNumberingTab: 0,
                    wBorderSpace: 0,
                    wBorderWidth: 0,
                    wBorders: 0,
                };

                SendMessageW(
                    edit_hwnd,
                    EM_SETPARAFORMAT,
                    0,
                    &pf as *const PARAFORMAT2 as isize,
                );

                // Create menu bar
                let hmenu = CreateMenu();

                // Create File menu
                let hmenu_file = CreateMenu();
                let new_text = format!("{}\0", get_string("MENU_NEW"));
                let new_text_utf16: Vec<u16> = new_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu_file,
                    MF_STRING,
                    ID_FILE_NEW as usize,
                    new_text_utf16.as_ptr(),
                );
                let open_text = format!("{}\0", get_string("MENU_OPEN"));
                let open_text_utf16: Vec<u16> = open_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu_file,
                    MF_STRING,
                    ID_FILE_OPEN as usize,
                    open_text_utf16.as_ptr(),
                );
                let save_text = format!("{}\0", get_string("MENU_SAVE"));
                let save_text_utf16: Vec<u16> = save_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu_file,
                    MF_STRING,
                    ID_FILE_SAVE as usize,
                    save_text_utf16.as_ptr(),
                );
                let saveas_text = format!("{}\0", get_string("MENU_SAVEAS"));
                let saveas_text_utf16: Vec<u16> = saveas_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu_file,
                    MF_STRING,
                    ID_FILE_SAVEAS as usize,
                    saveas_text_utf16.as_ptr(),
                );
                let exit_text = format!("{}\0", get_string("MENU_EXIT"));
                let exit_text_utf16: Vec<u16> = exit_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu_file,
                    MF_STRING,
                    ID_FILE_EXIT as usize,
                    exit_text_utf16.as_ptr(),
                );
                let file_text = format!("{}\0", get_string("MENU_FILE"));
                let file_text_utf16: Vec<u16> = file_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu,
                    MF_POPUP,
                    hmenu_file as usize,
                    file_text_utf16.as_ptr(),
                );

                // Create Edit menu
                let hmenu_edit = CreateMenu();
                let selectall_text = format!("{}\0", get_string("MENU_SELECTALL"));
                let selectall_text_utf16: Vec<u16> = selectall_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu_edit,
                    MF_STRING,
                    ID_EDIT_SELECTALL as usize,
                    selectall_text_utf16.as_ptr(),
                );
                let cut_text = format!("{}\0", get_string("MENU_CUT"));
                let cut_text_utf16: Vec<u16> = cut_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu_edit,
                    MF_STRING,
                    ID_EDIT_CUT as usize,
                    cut_text_utf16.as_ptr(),
                );
                let copy_text = format!("{}\0", get_string("MENU_COPY"));
                let copy_text_utf16: Vec<u16> = copy_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu_edit,
                    MF_STRING,
                    ID_EDIT_COPY as usize,
                    copy_text_utf16.as_ptr(),
                );
                let paste_text = format!("{}\0", get_string("MENU_PASTE"));
                let paste_text_utf16: Vec<u16> = paste_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu_edit,
                    MF_STRING,
                    ID_EDIT_PASTE as usize,
                    paste_text_utf16.as_ptr(),
                );
                let edit_text = format!("{}\0", get_string("MENU_EDIT"));
                let edit_text_utf16: Vec<u16> = edit_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu,
                    MF_POPUP,
                    hmenu_edit as usize,
                    edit_text_utf16.as_ptr(),
                );

                // Create View menu
                let hmenu_view = CreateMenu();
                let wordwrap_text = format!("{}\0", get_string("MENU_WORDWRAP"));
                let wordwrap_text_utf16: Vec<u16> = wordwrap_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu_view,
                    MF_STRING,
                    ID_VIEW_WORDWRAP as usize,
                    wordwrap_text_utf16.as_ptr(),
                );
                let statusbar_text = format!("{}\0", get_string("MENU_STATUSBAR"));
                let statusbar_text_utf16: Vec<u16> = statusbar_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu_view,
                    MF_STRING,
                    ID_VIEW_STATUSBAR as usize,
                    statusbar_text_utf16.as_ptr(),
                );
                let view_text = format!("{}\0", get_string("MENU_VIEW"));
                let view_text_utf16: Vec<u16> = view_text.encode_utf16().collect();
                AppendMenuW(
                    hmenu,
                    MF_POPUP,
                    hmenu_view as usize,
                    view_text_utf16.as_ptr(),
                );

                // Set menu
                SetMenu(hwnd, hmenu);

                // Store menu handle for later updates
                if let Ok(mut menu_handle) = MENU_HANDLE.lock() {
                    *menu_handle = Some(hmenu as isize);
                }

                // Set initial word wrap menu check state
                update_wordwrap_menu_check();

                // Set initial status bar menu check state
                update_statusbar_menu_check();

                // Create separator line above status bar using custom separator class
                let separator_class = "SeparatorClass\0".encode_utf16().collect::<Vec<_>>();
                let separator_hwnd = CreateWindowExW(
                    0,
                    separator_class.as_ptr(),
                    std::ptr::null(),
                    WS_CHILD | WS_VISIBLE,
                    0,
                    0,
                    0,
                    2,
                    hwnd,
                    std::ptr::null_mut(),
                    hinstance,
                    std::ptr::null_mut(),
                );
                SetWindowLongPtrW(hwnd, 16, separator_hwnd as isize);

                // Create 3 status bar sections
                let status_class = "StatusTextClass\0".encode_utf16().collect::<Vec<_>>();
                const SS_LEFT: u32 = 0x0000;
                const SS_RIGHT: u32 = 0x0002;
                const SS_CENTERIMAGE: u32 = 0x0200;

                // Character count (right-aligned text)
                let char_hwnd = CreateWindowExW(
                    0,
                    status_class.as_ptr(),
                    std::ptr::null(),
                    WS_CHILD | WS_VISIBLE | SS_RIGHT | SS_CENTERIMAGE,
                    0,
                    0,
                    0,
                    20,
                    hwnd,
                    std::ptr::null_mut(),
                    hinstance,
                    std::ptr::null_mut(),
                );

                // Line and column position
                let pos_hwnd = CreateWindowExW(
                    0,
                    status_class.as_ptr(),
                    std::ptr::null(),
                    WS_CHILD | WS_VISIBLE | SS_LEFT | SS_CENTERIMAGE,
                    0,
                    0,
                    0,
                    20,
                    hwnd,
                    std::ptr::null_mut(),
                    hinstance,
                    std::ptr::null_mut(),
                );

                // Encoding (UTF-8)
                let encoding_hwnd = CreateWindowExW(
                    0,
                    status_class.as_ptr(),
                    std::ptr::null(),
                    WS_CHILD | WS_VISIBLE | SS_LEFT | SS_CENTERIMAGE,
                    0,
                    0,
                    0,
                    20,
                    hwnd,
                    std::ptr::null_mut(),
                    hinstance,
                    std::ptr::null_mut(),
                );

                // Set font for status bars
                let font_name = "Segoe UI\0".encode_utf16().collect::<Vec<_>>();
                let hfont = CreateFontW(
                    -12, // 9pt (negative value for font height)
                    0,
                    0,
                    0,
                    400, // FW_NORMAL
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    font_name.as_ptr(),
                );
                SendMessageW(char_hwnd, WM_SETFONT, hfont as usize, 1);
                SendMessageW(pos_hwnd, WM_SETFONT, hfont as usize, 1);
                SendMessageW(encoding_hwnd, WM_SETFONT, hfont as usize, 1);

                // Set UTF-8 text
                let encoding_text = "UTF-8\0".encode_utf16().collect::<Vec<_>>();
                SetWindowTextW(encoding_hwnd, encoding_text.as_ptr());

                // Create vertical separator lines between sections using custom separator class
                let separator_class = "SeparatorClass\0".encode_utf16().collect::<Vec<_>>();
                let sep1_hwnd = CreateWindowExW(
                    0,
                    separator_class.as_ptr(),
                    std::ptr::null(),
                    WS_CHILD | WS_VISIBLE,
                    0,
                    0,
                    0,
                    20,
                    hwnd,
                    std::ptr::null_mut(),
                    hinstance,
                    std::ptr::null_mut(),
                );

                let sep2_hwnd = CreateWindowExW(
                    0,
                    separator_class.as_ptr(),
                    std::ptr::null(),
                    WS_CHILD | WS_VISIBLE,
                    0,
                    0,
                    0,
                    20,
                    hwnd,
                    std::ptr::null_mut(),
                    hinstance,
                    std::ptr::null_mut(),
                );

                // Store handles
                SetWindowLongPtrW(hwnd, 8, char_hwnd as isize);
                SetWindowLongPtrW(hwnd, 24, sep1_hwnd as isize);
                SetWindowLongPtrW(hwnd, 32, pos_hwnd as isize);
                SetWindowLongPtrW(hwnd, 40, sep2_hwnd as isize);
                SetWindowLongPtrW(hwnd, 48, encoding_hwnd as isize);

                // Set window icon from resources (ID 1)
                let hicon = LoadIconW(hinstance, 1 as *const u16);
                if hicon as usize != 0 {
                    SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon as isize);
                    SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon as isize);
                }

                0
            }
            WM_SIZE => {
                let edit_hwnd = GetWindowLongPtrW(hwnd, 0) as HWND;

                if edit_hwnd != HWND::default() {
                    let mut rect: RECT = std::mem::zeroed();
                    GetClientRect(hwnd, &mut rect);

                    // Check if status bar is visible
                    let is_statusbar_visible = if let Ok(visible) = STATUSBAR_VISIBLE.lock() {
                        *visible
                    } else {
                        true
                    };

                    let status_height = 24;
                    let separator_height = if is_statusbar_visible { 1 } else { 0 };
                    let status_total_height = if is_statusbar_visible {
                        status_height + separator_height
                    } else {
                        0
                    };
                    let edit_height = (rect.bottom - rect.top - status_total_height).max(0);

                    // Resize EDIT to account for status bar
                    SetWindowPos(
                        edit_hwnd,
                        std::ptr::null_mut(),
                        rect.left,
                        rect.top,
                        rect.right - rect.left,
                        edit_height,
                        SWP_NOZORDER,
                    );

                    // Position and resize separator line (full width including scrollbar)
                    let separator_hwnd = GetWindowLongPtrW(hwnd, 16) as HWND;
                    if separator_hwnd != HWND::default() {
                        SetWindowPos(
                            separator_hwnd,
                            std::ptr::null_mut(),
                            rect.left,
                            rect.top + edit_height,
                            rect.right - rect.left,
                            separator_height,
                            SWP_NOZORDER,
                        );
                    }

                    // Position and resize status bar sections at bottom (right-aligned)
                    let char_hwnd = GetWindowLongPtrW(hwnd, 8) as HWND;
                    let sep1_hwnd = GetWindowLongPtrW(hwnd, 24) as HWND;
                    let pos_hwnd = GetWindowLongPtrW(hwnd, 32) as HWND;
                    let sep2_hwnd = GetWindowLongPtrW(hwnd, 40) as HWND;
                    let encoding_hwnd = GetWindowLongPtrW(hwnd, 48) as HWND;
                    let scrollbar_width = 16;
                    let status_y = rect.top + edit_height + separator_height;

                    // Set widths for each section
                    let char_width = 80;
                    let separator_width = 2;
                    let pos_width = 120;
                    let encoding_width = 80;
                    let margin = 8;
                    let sep_margin = 8; // Margin around separators

                    let total_status_width = char_width
                        + (separator_width + sep_margin * 2)
                        + pos_width
                        + (separator_width + sep_margin * 2)
                        + encoding_width;
                    let start_x = rect.right - scrollbar_width - total_status_width - margin;

                    if char_hwnd != HWND::default() {
                        SetWindowPos(
                            char_hwnd,
                            std::ptr::null_mut(),
                            start_x,
                            status_y,
                            char_width,
                            status_height,
                            SWP_NOZORDER,
                        );
                    }

                    if sep1_hwnd != HWND::default() {
                        SetWindowPos(
                            sep1_hwnd,
                            std::ptr::null_mut(),
                            start_x + char_width + sep_margin,
                            status_y,
                            separator_width,
                            status_height,
                            SWP_NOZORDER,
                        );
                    }

                    if pos_hwnd != HWND::default() {
                        SetWindowPos(
                            pos_hwnd,
                            std::ptr::null_mut(),
                            start_x + char_width + sep_margin + separator_width + sep_margin,
                            status_y,
                            pos_width,
                            status_height,
                            SWP_NOZORDER,
                        );
                    }

                    if sep2_hwnd != HWND::default() {
                        SetWindowPos(
                            sep2_hwnd,
                            std::ptr::null_mut(),
                            start_x
                                + char_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + pos_width
                                + sep_margin,
                            status_y,
                            separator_width,
                            status_height,
                            SWP_NOZORDER,
                        );
                    }

                    if encoding_hwnd != HWND::default() {
                        SetWindowPos(
                            encoding_hwnd,
                            std::ptr::null_mut(),
                            start_x
                                + char_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + pos_width
                                + sep_margin
                                + separator_width
                                + sep_margin,
                            status_y,
                            encoding_width,
                            status_height,
                            SWP_NOZORDER,
                        );
                    }
                }

                0
            }
            WM_NOTIFY => {
                // Handle RichEdit notifications to remove OLE objects
                let edit_hwnd = GetWindowLongPtrW(hwnd, 0) as HWND;

                // Remove any OLE objects that were pasted/inserted
                remove_ole_objects(edit_hwnd);

                0
            }
            WM_PASTE => {
                // Handle paste - insert text only from clipboard
                let edit_hwnd = GetWindowLongPtrW(hwnd, 0) as HWND;

                if OpenClipboard(hwnd) != 0 {
                    let hdata = GetClipboardData(13); // CF_UNICODETEXT = 13

                    if !hdata.is_null() {
                        let text_ptr = hdata as *const u16;
                        let mut len = 0;

                        // Find the length of the string
                        while *text_ptr.add(len) != 0 {
                            len += 1;
                        }

                        // Convert to Rust string and insert into RichEdit
                        if len > 0 {
                            let text_slice = std::slice::from_raw_parts(text_ptr, len);
                            if let Ok(text) = String::from_utf16(text_slice) {
                                // Replace selected text with clipboard text
                                let text_utf16: Vec<u16> = text.encode_utf16().collect();
                                SendMessageW(edit_hwnd, 0x00C2, 1, text_utf16.as_ptr() as isize); // EM_REPLACESEL
                            }
                        }
                    }

                    CloseClipboard();
                }

                0
            }
            WM_SETCURSOR => {
                let hit_test = (lparam & 0xFFFF) as u32;
                const HTCLIENT: u32 = 1;

                // Only handle when in client area
                if hit_test == HTCLIENT {
                    let cursor_hwnd = wparam as HWND;

                    // Check if status bar is visible
                    let is_statusbar_visible = if let Ok(visible) = STATUSBAR_VISIBLE.lock() {
                        *visible
                    } else {
                        true
                    };

                    if is_statusbar_visible {
                        // Get cursor position
                        let mut cursor_pt = std::mem::zeroed();
                        GetCursorPos(&mut cursor_pt);

                        let mut window_rect: RECT = std::mem::zeroed();
                        GetWindowRect(hwnd, &mut window_rect);

                        let mut client_rect: RECT = std::mem::zeroed();
                        GetClientRect(hwnd, &mut client_rect);

                        // Calculate status bar area
                        let status_height = 24;
                        let separator_height = 1;
                        let status_total_height = status_height + separator_height;

                        // Calculate status bar top position in screen coordinates
                        let status_bar_top = window_rect.bottom - status_total_height;

                        // If cursor is in status bar area (including empty spaces), set arrow cursor
                        if cursor_pt.y >= status_bar_top {
                            let cursor = LoadCursorW(std::ptr::null_mut(), IDC_ARROW);
                            SetCursor(cursor);
                            return 1;
                        }
                    }

                    // Also handle child window status bar components
                    if cursor_hwnd != hwnd {
                        let char_hwnd = GetWindowLongPtrW(hwnd, 8) as HWND;
                        let sep1_hwnd = GetWindowLongPtrW(hwnd, 24) as HWND;
                        let pos_hwnd = GetWindowLongPtrW(hwnd, 32) as HWND;
                        let sep2_hwnd = GetWindowLongPtrW(hwnd, 40) as HWND;
                        let encoding_hwnd = GetWindowLongPtrW(hwnd, 48) as HWND;
                        let separator_hwnd = GetWindowLongPtrW(hwnd, 16) as HWND;

                        if cursor_hwnd == char_hwnd
                            || cursor_hwnd == sep1_hwnd
                            || cursor_hwnd == pos_hwnd
                            || cursor_hwnd == sep2_hwnd
                            || cursor_hwnd == encoding_hwnd
                            || cursor_hwnd == separator_hwnd
                        {
                            let cursor = LoadCursorW(std::ptr::null_mut(), IDC_ARROW);
                            SetCursor(cursor);
                            return 1;
                        }
                    }
                }

                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CONTEXTMENU => {
                // Show context menu when right-clicked in RichEdit
                let mut pt = std::mem::zeroed();
                GetCursorPos(&mut pt);
                show_context_menu(hwnd, pt.x, pt.y);
                0
            }
            WM_COMMAND => {
                let edit_hwnd = GetWindowLongPtrW(hwnd, 0) as HWND;
                let status_hwnd = GetWindowLongPtrW(hwnd, 8) as HWND;
                let cmd_id = wparam as i32;

                match cmd_id {
                    ID_FILE_NEW => {
                        // Clear the editor with empty string
                        let empty = "\0".encode_utf16().collect::<Vec<_>>();
                        SendMessageW(edit_hwnd, 0x000C, 0, empty.as_ptr() as isize); // WM_SETTEXT

                        // Reset to "Untitled" or "無題"
                        let default_filename = get_string("FILE_UNTITLED");
                        let untitled_path = std::path::PathBuf::from(default_filename);
                        if let Ok(mut current_file) = CURRENT_FILE.lock() {
                            *current_file = Some(untitled_path);
                        }

                        // Reset modified flag
                        SendMessageW(edit_hwnd, EM_SETMODIFY, 0, 0);

                        0
                    }
                    ID_FILE_OPEN => {
                        if let Some(path) = file_io::open_file_dialog() {
                            if let Ok(content) = file_io::load_file(&path) {
                                // Convert UTF-8 to UTF-16 with null terminator
                                let utf16: Vec<u16> =
                                    content.encode_utf16().chain(std::iter::once(0)).collect();

                                // Use WM_SETTEXT with UTF-16 string
                                SendMessageW(edit_hwnd, 0x000C, 0, utf16.as_ptr() as isize);

                                // Store the file path
                                if let Ok(mut current_file) = CURRENT_FILE.lock() {
                                    *current_file = Some(path.clone());
                                }

                                // Reset modified flag
                                SendMessageW(edit_hwnd, EM_SETMODIFY, 0, 0);

                                // Update title with filename
                                if let Some(filename) = path.file_name() {
                                    if let Some(filename_str) = filename.to_str() {
                                        let app_name = get_string("WINDOW_TITLE");
                                        let title = format!("{} - {}\0", filename_str, app_name);
                                        let title_utf16: Vec<u16> = title.encode_utf16().collect();
                                        SetWindowTextW(hwnd, title_utf16.as_ptr());
                                    }
                                }
                            }
                        }
                        0
                    }
                    ID_FILE_SAVE => {
                        if let Ok(current_file) = CURRENT_FILE.lock() {
                            if let Some(path) = current_file.as_ref() {
                                // Check if this is an untitled file
                                if is_untitled_file(path) {
                                    // For untitled files, use "Save As" dialog
                                    drop(current_file); // Release lock before showing dialog
                                    if let Some(new_path) = file_io::save_file_dialog() {
                                        let text_len =
                                            SendMessageW(edit_hwnd, 0x000E, 0, 0) as usize;
                                        if text_len > 0 {
                                            let mut buffer: Vec<u16> = vec![0; text_len + 1];
                                            SendMessageW(
                                                edit_hwnd,
                                                0x000D,
                                                (text_len + 1) as usize,
                                                buffer.as_mut_ptr() as isize,
                                            );
                                            if let Ok(text) =
                                                String::from_utf16(&buffer[..text_len])
                                            {
                                                let _ = file_io::save_file(&new_path, &text);
                                                *CURRENT_FILE.lock().unwrap() =
                                                    Some(new_path.clone());
                                                SendMessageW(edit_hwnd, EM_SETMODIFY, 0, 0);
                                                if let Some(filename) = new_path.file_name() {
                                                    if let Some(filename_str) = filename.to_str() {
                                                        let app_name = get_string("WINDOW_TITLE");
                                                        let title = format!(
                                                            "{} - {}\0",
                                                            filename_str, app_name
                                                        );
                                                        let title_utf16: Vec<u16> =
                                                            title.encode_utf16().collect();
                                                        SetWindowTextW(hwnd, title_utf16.as_ptr());
                                                    }
                                                }
                                            }
                                        }
                                    }
                                } else {
                                    // Regular save
                                    let text_len = SendMessageW(edit_hwnd, 0x000E, 0, 0) as usize;
                                    if text_len > 0 {
                                        let mut buffer: Vec<u16> = vec![0; text_len + 1];
                                        SendMessageW(
                                            edit_hwnd,
                                            0x000D,
                                            (text_len + 1) as usize,
                                            buffer.as_mut_ptr() as isize,
                                        );
                                        if let Ok(text) = String::from_utf16(&buffer[..text_len]) {
                                            let _ = file_io::save_file(path, &text);
                                            SendMessageW(edit_hwnd, EM_SETMODIFY, 0, 0);
                                            if let Some(filename) = path.file_name() {
                                                if let Some(filename_str) = filename.to_str() {
                                                    let app_name = get_string("WINDOW_TITLE");
                                                    let title = format!(
                                                        "{} - {}\0",
                                                        filename_str, app_name
                                                    );
                                                    let title_utf16: Vec<u16> =
                                                        title.encode_utf16().collect();
                                                    SetWindowTextW(hwnd, title_utf16.as_ptr());
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        0
                    }
                    ID_FILE_SAVEAS => {
                        // Show save dialog
                        if let Some(new_path) = file_io::save_file_dialog() {
                            let text_len = SendMessageW(edit_hwnd, 0x000E, 0, 0) as usize;
                            if text_len > 0 {
                                let mut buffer: Vec<u16> = vec![0; text_len + 1];
                                SendMessageW(
                                    edit_hwnd,
                                    0x000D,
                                    (text_len + 1) as usize,
                                    buffer.as_mut_ptr() as isize,
                                );
                                if let Ok(text) = String::from_utf16(&buffer[..text_len]) {
                                    let _ = file_io::save_file(&new_path, &text);
                                    *CURRENT_FILE.lock().unwrap() = Some(new_path.clone());
                                    SendMessageW(edit_hwnd, EM_SETMODIFY, 0, 0);
                                    if let Some(filename) = new_path.file_name() {
                                        if let Some(filename_str) = filename.to_str() {
                                            let app_name = get_string("WINDOW_TITLE");
                                            let title =
                                                format!("{} - {}\0", filename_str, app_name);
                                            let title_utf16: Vec<u16> =
                                                title.encode_utf16().collect();
                                            SetWindowTextW(hwnd, title_utf16.as_ptr());
                                        }
                                    }
                                }
                            }
                        }
                        0
                    }
                    ID_FILE_EXIT => {
                        PostQuitMessage(0);
                        0
                    }
                    ID_EDIT_SELECTALL => {
                        SendMessageW(edit_hwnd, EM_SETSEL as u32, 0, -1 as isize as isize);
                        // Update status bar
                        if status_hwnd != HWND::default() {
                            let mut start: usize = 0;
                            SendMessageW(
                                edit_hwnd,
                                EM_GETSEL,
                                &mut start as *mut usize as usize,
                                &mut start as *mut usize as isize,
                            );
                            let line = SendMessageW(edit_hwnd, EM_LINEFROMCHAR, start, 0) as i32;
                            let line_start =
                                SendMessageW(edit_hwnd, 0x00C1, line as usize, 0) as usize;
                            let col = (start - line_start) as i32;
                            let status_format = get_string("STATUS_LINE_COL");
                            let text = format!(
                                "{}\0",
                                status_format
                                    .replace("{line}", &(line + 1).to_string())
                                    .replace("{col}", &(col + 1).to_string())
                            );
                            let text_utf16: Vec<u16> = text.encode_utf16().collect();
                            SetWindowTextW(status_hwnd, text_utf16.as_ptr());
                        }
                        0
                    }
                    ID_EDIT_CUT => {
                        SendMessageW(edit_hwnd, WM_CUT, 0, 0);
                        0
                    }
                    ID_EDIT_COPY => {
                        SendMessageW(edit_hwnd, WM_COPY, 0, 0);
                        0
                    }
                    ID_EDIT_PASTE => {
                        // Send WM_PASTE message to trigger our custom paste handler
                        SendMessageW(hwnd, WM_PASTE, 0, 0);
                        0
                    }
                    ID_VIEW_WORDWRAP => {
                        toggle_word_wrap(edit_hwnd);
                        0
                    }
                    ID_VIEW_STATUSBAR => {
                        // Toggle status bar visibility
                        let new_visibility = if let Ok(mut visible) = STATUSBAR_VISIBLE.lock() {
                            *visible = !*visible;
                            *visible
                        } else {
                            true
                        };

                        // Get status bar handles
                        let char_hwnd = GetWindowLongPtrW(hwnd, 8) as HWND;
                        let sep1_hwnd = GetWindowLongPtrW(hwnd, 24) as HWND;
                        let pos_hwnd = GetWindowLongPtrW(hwnd, 32) as HWND;
                        let sep2_hwnd = GetWindowLongPtrW(hwnd, 40) as HWND;
                        let encoding_hwnd = GetWindowLongPtrW(hwnd, 48) as HWND;
                        let separator_hwnd = GetWindowLongPtrW(hwnd, 16) as HWND;

                        // Show/hide all status bar windows completely
                        let show_cmd = if new_visibility { 5 } else { 0 }; // SW_SHOW = 5, SW_HIDE = 0
                        ShowWindow(char_hwnd, show_cmd);
                        ShowWindow(sep1_hwnd, show_cmd);
                        ShowWindow(pos_hwnd, show_cmd);
                        ShowWindow(sep2_hwnd, show_cmd);
                        ShowWindow(encoding_hwnd, show_cmd);
                        ShowWindow(separator_hwnd, show_cmd);

                        // Update menu check state
                        if let Ok(menu_handle) = MENU_HANDLE.lock() {
                            if let Some(hmenu_isize) = *menu_handle {
                                let check_state = if new_visibility {
                                    MF_CHECKED
                                } else {
                                    MF_UNCHECKED
                                };
                                CheckMenuItem(
                                    hmenu_isize as *mut std::ffi::c_void,
                                    ID_VIEW_STATUSBAR as u32,
                                    check_state,
                                );
                            }
                        }

                        // Trigger WM_SIZE to recalculate layout without moving window position
                        let mut rect: RECT = std::mem::zeroed();
                        GetClientRect(hwnd, &mut rect);
                        let width = rect.right - rect.left;
                        let height = rect.bottom - rect.top;
                        SendMessageW(
                            hwnd,
                            WM_SIZE,
                            0,
                            ((height as isize) << 16) | (width as isize),
                        );

                        // Invalidate and redraw
                        InvalidateRect(hwnd, std::ptr::null(), 1);

                        0
                    }
                    _ => DefWindowProcW(hwnd, msg, wparam, lparam),
                }
            }
            0x0101 | 0x0202 => {
                // WM_KEYUP | WM_LBUTTONUP
                // Only update status bar if it's visible
                if let Ok(visible) = STATUSBAR_VISIBLE.lock() {
                    if *visible {
                        let edit_hwnd = GetWindowLongPtrW(hwnd, 0) as HWND;
                        let char_hwnd = GetWindowLongPtrW(hwnd, 8) as HWND;
                        let pos_hwnd = GetWindowLongPtrW(hwnd, 32) as HWND;
                        update_status_bar(edit_hwnd, char_hwnd, pos_hwnd);
                    }
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CLOSE => {
                DestroyWindow(hwnd);
                0
            }
            WM_DESTROY => {
                PostQuitMessage(0);
                0
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

fn main() {
    // Initialize language based on system locale
    init_language();

    // Initialize with "Untitled" or "無題" file based on language
    let default_filename = get_string("FILE_UNTITLED");
    let untitled_path = std::path::PathBuf::from(default_filename);
    *CURRENT_FILE.lock().unwrap() = Some(untitled_path);

    unsafe {
        let hinstance = GetModuleHandleW(std::ptr::null());
        let class_name = "NotepadWindowClass\0".encode_utf16().collect::<Vec<_>>();

        const COLOR_BTNFACE: i32 = 15;

        // Load the icon from resources
        let hicon = LoadIconW(hinstance, std::ptr::null());

        let wnd_class = WNDCLASSW {
            style: CS_VREDRAW | CS_HREDRAW,
            lpfnWndProc: Some(window_proc),
            cbClsExtra: 0,
            cbWndExtra: (std::mem::size_of::<isize>() * 8) as i32, // 8 pointers: EDIT, separator, char, sep1, line, sep2, col
            hInstance: hinstance,
            hIcon: hicon,
            hCursor: std::ptr::null_mut(),
            hbrBackground: GetSysColorBrush(COLOR_BTNFACE),
            lpszMenuName: std::ptr::null(),
            lpszClassName: class_name.as_ptr(),
        };

        RegisterClassW(&wnd_class);

        // Register status bar window classes (separator and text)
        status_bar::register_status_bar_classes();

        let window_title_str = format!("{}\0", get_string("WINDOW_TITLE"));
        let window_title = window_title_str.encode_utf16().collect::<Vec<_>>();

        const WS_THICKFRAME: u32 = 0x00040000;
        let hwnd = CreateWindowExW(
            0,
            class_name.as_ptr(),
            window_title.as_ptr(),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE | WS_THICKFRAME,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            800,
            600,
            std::ptr::null_mut(),
            std::ptr::null_mut(),
            hinstance,
            std::ptr::null_mut(),
        );

        // Set initial title with "Untitled" or "無題" based on language
        let app_name = get_string("WINDOW_TITLE");
        let default_filename = get_string("FILE_UNTITLED");
        let initial_title = format!("{} - {}\0", default_filename, app_name);
        let initial_title_utf16: Vec<u16> = initial_title.encode_utf16().collect();
        SetWindowTextW(hwnd, initial_title_utf16.as_ptr());

        let mut msg: MSG = std::mem::zeroed();
        while GetMessageW(&mut msg, std::ptr::null_mut(), 0, 0) > 0 {
            // Block Ctrl+E and Ctrl+R to prevent editor glitches
            if msg.message == WM_KEYDOWN && (msg.wParam as i32 == 0x45 || msg.wParam as i32 == 0x52)
            {
                let ctrl_pressed = (GetKeyState(0x11) as u16 & 0x8000) != 0; // VK_CONTROL = 0x11
                if ctrl_pressed {
                    continue; // Skip Ctrl+E and Ctrl+R
                }
            }

            // Check for Ctrl+S before dispatching
            if msg.message == WM_KEYDOWN && msg.wParam as i32 == 0x53 {
                let ctrl_pressed = (GetKeyState(0x11) as u16 & 0x8000) != 0; // VK_CONTROL = 0x11
                if ctrl_pressed {
                    let edit_hwnd = GetWindowLongPtrW(hwnd, 0) as HWND;
                    // Trigger save
                    if let Ok(current_file) = CURRENT_FILE.lock() {
                        if let Some(path) = current_file.as_ref() {
                            // Check if this is an untitled file
                            if is_untitled_file(path) {
                                // For untitled files, use "Save As" dialog
                                drop(current_file); // Release lock before showing dialog
                                if let Some(new_path) = file_io::save_file_dialog() {
                                    let text_len = SendMessageW(edit_hwnd, 0x000E, 0, 0) as usize;
                                    if text_len > 0 {
                                        let mut buffer: Vec<u16> = vec![0; text_len + 1];
                                        SendMessageW(
                                            edit_hwnd,
                                            0x000D,
                                            (text_len + 1) as usize,
                                            buffer.as_mut_ptr() as isize,
                                        );
                                        if let Ok(text) = String::from_utf16(&buffer[..text_len]) {
                                            let _ = file_io::save_file(&new_path, &text);
                                            *CURRENT_FILE.lock().unwrap() = Some(new_path.clone());
                                            SendMessageW(edit_hwnd, EM_SETMODIFY, 0, 0);
                                            if let Some(filename) = new_path.file_name() {
                                                if let Some(filename_str) = filename.to_str() {
                                                    let app_name = get_string("WINDOW_TITLE");
                                                    let title = format!(
                                                        "{} - {}\0",
                                                        filename_str, app_name
                                                    );
                                                    let title_utf16: Vec<u16> =
                                                        title.encode_utf16().collect();
                                                    SetWindowTextW(hwnd, title_utf16.as_ptr());
                                                }
                                            }
                                        }
                                    }
                                }
                            } else {
                                // Regular save
                                let text_len = SendMessageW(edit_hwnd, 0x000E, 0, 0) as usize;
                                if text_len > 0 {
                                    let mut buffer: Vec<u16> = vec![0; text_len + 1];
                                    SendMessageW(
                                        edit_hwnd,
                                        0x000D,
                                        (text_len + 1) as usize,
                                        buffer.as_mut_ptr() as isize,
                                    );
                                    if let Ok(text) = String::from_utf16(&buffer[..text_len]) {
                                        let _ = file_io::save_file(path, &text);
                                        SendMessageW(edit_hwnd, EM_SETMODIFY, 0, 0);
                                        if let Some(filename) = path.file_name() {
                                            if let Some(filename_str) = filename.to_str() {
                                                let app_name = get_string("WINDOW_TITLE");
                                                let title =
                                                    format!("{} - {}\0", filename_str, app_name);
                                                let title_utf16: Vec<u16> =
                                                    title.encode_utf16().collect();
                                                SetWindowTextW(hwnd, title_utf16.as_ptr());
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    continue; // Skip default processing
                }
            }

            // Check for Ctrl+V before dispatching
            if msg.message == WM_KEYDOWN && msg.wParam as i32 == 0x56 {
                let ctrl_pressed = (GetKeyState(0x11) as u16 & 0x8000) != 0; // VK_CONTROL = 0x11
                if ctrl_pressed {
                    let edit_hwnd = GetWindowLongPtrW(hwnd, 0) as HWND;

                    // Custom paste handler: extract text from clipboard only
                    if OpenClipboard(hwnd) != 0 {
                        let hdata = GetClipboardData(13); // CF_UNICODETEXT = 13

                        if !hdata.is_null() {
                            let text_ptr = hdata as *const u16;
                            let mut len = 0;

                            // Find the length of the string
                            while *text_ptr.add(len) != 0 {
                                len += 1;
                            }

                            // Convert to Rust string and insert into RichEdit
                            if len > 0 {
                                let text_slice = std::slice::from_raw_parts(text_ptr, len);
                                if let Ok(text) = String::from_utf16(text_slice) {
                                    // Replace selected text with clipboard text
                                    let text_utf16: Vec<u16> =
                                        text.encode_utf16().chain(std::iter::once(0)).collect();
                                    SendMessageW(
                                        edit_hwnd,
                                        0x00C2,
                                        1,
                                        text_utf16.as_ptr() as isize,
                                    ); // EM_REPLACESEL
                                }
                            }
                        }

                        CloseClipboard();
                    }

                    continue; // Skip default processing - don't let RichEdit handle Ctrl+V
                }
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);

            // Always update status bar and title after every message
            let edit_hwnd = GetWindowLongPtrW(hwnd, 0) as HWND;

            // Only update status bar if it's visible
            if let Ok(visible) = STATUSBAR_VISIBLE.lock() {
                if *visible {
                    let char_hwnd = GetWindowLongPtrW(hwnd, 8) as HWND;
                    let pos_hwnd = GetWindowLongPtrW(hwnd, 32) as HWND;
                    status_bar::update_status_bar(edit_hwnd, char_hwnd, pos_hwnd);
                }
            }

            update_title_if_needed(hwnd, edit_hwnd);
        }
    }
}
