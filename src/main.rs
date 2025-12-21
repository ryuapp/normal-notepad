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
use file_io::FileEncoding;
use i18n::{get_string, init_language};
use status_bar::update_status_bar;
use std::path::PathBuf;
use std::sync::Mutex;

// Global variables for file state
static CURRENT_FILE: Mutex<Option<PathBuf>> = Mutex::new(None);
static LAST_MODIFIED_STATE: Mutex<bool> = Mutex::new(false);
static SAVED_CONTENT: Mutex<String> = Mutex::new(String::new());
static WORD_WRAP_ENABLED: Mutex<bool> = Mutex::new(true);
static STATUSBAR_VISIBLE: Mutex<bool> = Mutex::new(true);
static MENU_HANDLE: Mutex<Option<isize>> = Mutex::new(None);
static CURRENT_ENCODING: Mutex<FileEncoding> = Mutex::new(FileEncoding::Utf8);

use windows::Win32::Foundation::HINSTANCE;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    CreateFontW, FONT_CHARSET, FONT_CLIP_PRECISION, FONT_OUTPUT_PRECISION, FONT_QUALITY,
    GetSysColorBrush, HBRUSH, InvalidateRect, SYS_COLOR_INDEX,
};
use windows::Win32::System::DataExchange::{CloseClipboard, GetClipboardData, OpenClipboard};
use windows::Win32::System::LibraryLoader::{GetModuleHandleW, LoadLibraryW};
use windows::Win32::UI::Controls::{EM_SETMARGINS, EM_SETMODIFY, EM_SETSEL};
use windows::Win32::UI::Input::KeyboardAndMouse::GetKeyState;
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CheckMenuItem, CreateMenu, CreateWindowExW, DefWindowProcW, DestroyWindow,
    DispatchMessageW, GetClientRect, GetCursorPos, GetMessageW, GetWindowLongPtrW, GetWindowRect,
    HMENU, IDC_ARROW, LoadCursorW, LoadIconW, MENU_ITEM_FLAGS, MSG, PostQuitMessage,
    RegisterClassW, SET_WINDOW_POS_FLAGS, SHOW_WINDOW_CMD, SendMessageW, SetCursor, SetMenu,
    SetWindowLongPtrW, SetWindowPos, SetWindowTextW, ShowWindow, TranslateMessage, WINDOW_EX_STYLE,
    WINDOW_LONG_PTR_INDEX, WINDOW_STYLE, WM_CLOSE, WM_COMMAND, WM_CONTEXTMENU, WM_COPY, WM_CREATE,
    WM_CUT, WM_DESTROY, WM_GETMINMAXINFO, WM_KEYDOWN, WM_NOTIFY, WM_PASTE, WM_SETCURSOR,
    WM_SETFONT, WM_SETICON, WM_SIZE, WNDCLASS_STYLES, WNDCLASSW,
};
use windows::core::PCWSTR;

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
        // Get current text content
        let text_len = SendMessageW(edit_hwnd, 0x000E, Some(WPARAM(0)), Some(LPARAM(0))).0 as usize; // WM_GETTEXTLENGTH
        let current_text = if text_len > 0 {
            let mut buffer: Vec<u16> = vec![0; text_len + 1];
            SendMessageW(
                edit_hwnd,
                0x000D,
                Some(WPARAM((text_len + 1) as usize)),
                Some(LPARAM(buffer.as_mut_ptr() as isize)),
            );
            String::from_utf16(&buffer[..text_len]).unwrap_or_default()
        } else {
            String::new()
        };

        // Compare with saved content
        let is_modified = if let Ok(saved) = SAVED_CONTENT.lock() {
            *saved != current_text
        } else {
            false
        };

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
                        let _ = SetWindowTextW(hwnd, PCWSTR(title_utf16.as_ptr()));
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
        let text_length =
            SendMessageW(edit_hwnd, 0x000E, Some(WPARAM(0)), Some(LPARAM(0))).0 as i32; // WM_GETTEXTLENGTH

        if text_length <= 0 {
            return;
        }

        // Allocate buffer and get text
        let mut buffer = vec![0u16; (text_length + 1) as usize];
        let actual_len = SendMessageW(
            edit_hwnd,
            EM_GETTEXT,
            Some(WPARAM((text_length + 1) as usize)),
            Some(LPARAM(buffer.as_mut_ptr() as isize)),
        )
        .0 as usize;

        if actual_len == 0 {
            return;
        }

        // Check for OLE placeholder characters
        let mut has_ole = false;
        buffer.retain(|&ch| {
            if ch == 0xFFFC || ch == OLE_PLACEHOLDER || ch == 0xFFFD {
                has_ole = true;
                false
            } else {
                true
            }
        });

        // If OLE objects were found and removed, update the text
        if has_ole {
            let new_text: Vec<u16> = buffer.iter().copied().chain(std::iter::once(0)).collect();
            SendMessageW(
                edit_hwnd,
                EM_SETTEXT,
                Some(WPARAM(0)),
                Some(LPARAM(new_text.as_ptr() as isize)),
            );
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
            SendMessageW(
                edit_hwnd,
                EM_SETTARGETDEVICE,
                Some(WPARAM(0)),
                Some(LPARAM(0)),
            );
        } else {
            SendMessageW(
                edit_hwnd,
                EM_SETTARGETDEVICE,
                Some(WPARAM(0)),
                Some(LPARAM(1)),
            );
        }
    }

    update_wordwrap_menu_check();
}

// Helper function to update word wrap menu check state
fn update_wordwrap_menu_check() {
    if let Ok(menu_handle) = MENU_HANDLE.lock() {
        if let Some(hmenu_isize) = *menu_handle {
            if let Ok(wrap_state) = WORD_WRAP_ENABLED.lock() {
                let check_state = if *wrap_state {
                    MENU_ITEM_FLAGS(0x00000008) // MF_CHECKED
                } else {
                    MENU_ITEM_FLAGS(0x00000000) // MF_UNCHECKED
                };
                unsafe {
                    let _ = CheckMenuItem(
                        HMENU(hmenu_isize as *mut core::ffi::c_void),
                        ID_VIEW_WORDWRAP as u32,
                        check_state.0,
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
                let check_state = if *visible {
                    MENU_ITEM_FLAGS(0x00000008)
                } else {
                    MENU_ITEM_FLAGS(0x00000000)
                };
                unsafe {
                    let _ = CheckMenuItem(
                        HMENU(hmenu_isize as *mut core::ffi::c_void),
                        ID_VIEW_STATUSBAR as u32,
                        check_state.0,
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
                let hinstance = GetModuleHandleW(None).unwrap_or_default();

                // Load RichEdit library
                let richedit_lib = "Msftedit.dll\0".encode_utf16().collect::<Vec<_>>();
                let _ = LoadLibraryW(PCWSTR(richedit_lib.as_ptr()));

                let richedit_class = "RICHEDIT50W\0".encode_utf16().collect::<Vec<_>>();

                let edit_hwnd = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(richedit_class.as_ptr()),
                    PCWSTR::null(),
                    WINDOW_STYLE(0x40000000 | 0x10000000 | 0x00200000 | 0x00100000 | ES_MULTILINE), // WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL
                    0,
                    0,
                    0,
                    0,
                    Some(hwnd),
                    None,
                    Some(HINSTANCE(hinstance.0)),
                    None,
                )
                .unwrap_or_default();

                // Store the RichEdit control handle
                SetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(0), edit_hwnd.0 as isize);

                // Set unlimited text size
                SendMessageW(
                    edit_hwnd,
                    EM_EXLIMITTEXT,
                    Some(WPARAM(0)),
                    Some(LPARAM(u64::MAX as isize)),
                );

                // Enable word wrap by default
                SendMessageW(
                    edit_hwnd,
                    EM_SETTARGETDEVICE,
                    Some(WPARAM(0)),
                    Some(LPARAM(0)),
                );

                // Set margins
                const EC_LEFTMARGIN: u32 = 0x0001;
                let margin = (8u32) | ((0u32) << 16);
                SendMessageW(
                    edit_hwnd,
                    EM_SETMARGINS,
                    Some(WPARAM(EC_LEFTMARGIN as usize)),
                    Some(LPARAM(margin as isize)),
                );

                let margin = (8u32) | ((0u32) << 16);
                SendMessageW(
                    edit_hwnd,
                    EM_SETMARGINS,
                    Some(WPARAM(EC_TOPMARGIN as usize)),
                    Some(LPARAM(margin as isize)),
                );

                // Set default font
                let font_name = "MS Gothic";
                let hfont_edit = CreateFontW(
                    -16,                      // cHeight
                    0,                        // cWidth
                    0,                        // cEscapement
                    0,                        // cOrientation
                    400,                      // cWeight (FW_NORMAL)
                    0,                        // bItalic
                    0,                        // bUnderline
                    0,                        // bStrikeOut
                    FONT_CHARSET(0),          // iCharSet
                    FONT_OUTPUT_PRECISION(0), // iOutPrecision
                    FONT_CLIP_PRECISION(0),   // iClipPrecision
                    FONT_QUALITY(0),          // iQuality
                    0,                        // iPitchAndFamily
                    windows::core::PCWSTR(
                        font_name
                            .encode_utf16()
                            .chain(Some(0))
                            .collect::<Vec<_>>()
                            .as_ptr(),
                    ),
                );
                SendMessageW(
                    edit_hwnd,
                    WM_SETFONT,
                    Some(WPARAM(hfont_edit.0 as usize)),
                    Some(LPARAM(1)),
                );

                // Disable auto font
                let lang_options = SendMessageW(
                    edit_hwnd,
                    EM_GETLANGOPTIONS,
                    Some(WPARAM(0)),
                    Some(LPARAM(0)),
                );
                let new_options = LRESULT(lang_options.0 & !(IMF_AUTOFONT | IMF_DUALFONT) as isize);
                SendMessageW(
                    edit_hwnd,
                    EM_SETLANGOPTIONS,
                    Some(WPARAM(0)),
                    Some(LPARAM(new_options.0)),
                );

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
                    dyLineSpacing: 260,
                    sStyle: 0,
                    bLineSpacingRule: 4,
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
                    Some(WPARAM(0)),
                    Some(LPARAM(&pf as *const PARAFORMAT2 as isize)),
                );

                // Create menu bar
                let hmenu = CreateMenu().unwrap_or_default();

                // Create File menu
                let hmenu_file = CreateMenu().unwrap_or_default();
                let new_text = format!("{}\0", get_string("MENU_NEW"));
                let new_text_utf16: Vec<u16> = new_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu_file,
                    MENU_ITEM_FLAGS(0x00000000), // MF_STRING
                    ID_FILE_NEW as usize,
                    PCWSTR(new_text_utf16.as_ptr()),
                );
                let open_text = format!("{}\0", get_string("MENU_OPEN"));
                let open_text_utf16: Vec<u16> = open_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu_file,
                    MENU_ITEM_FLAGS(0x00000000),
                    ID_FILE_OPEN as usize,
                    PCWSTR(open_text_utf16.as_ptr()),
                );
                let save_text = format!("{}\0", get_string("MENU_SAVE"));
                let save_text_utf16: Vec<u16> = save_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu_file,
                    MENU_ITEM_FLAGS(0x00000000),
                    ID_FILE_SAVE as usize,
                    PCWSTR(save_text_utf16.as_ptr()),
                );
                let saveas_text = format!("{}\0", get_string("MENU_SAVEAS"));
                let saveas_text_utf16: Vec<u16> = saveas_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu_file,
                    MENU_ITEM_FLAGS(0x00000000),
                    ID_FILE_SAVEAS as usize,
                    PCWSTR(saveas_text_utf16.as_ptr()),
                );
                let exit_text = format!("{}\0", get_string("MENU_EXIT"));
                let exit_text_utf16: Vec<u16> = exit_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu_file,
                    MENU_ITEM_FLAGS(0x00000000),
                    ID_FILE_EXIT as usize,
                    PCWSTR(exit_text_utf16.as_ptr()),
                );
                let file_text = format!("{}\0", get_string("MENU_FILE"));
                let file_text_utf16: Vec<u16> = file_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu,
                    MENU_ITEM_FLAGS(0x00000010), // MF_POPUP
                    hmenu_file.0 as usize,
                    PCWSTR(file_text_utf16.as_ptr()),
                );

                // Create Edit menu
                let hmenu_edit = CreateMenu().unwrap_or_default();
                let selectall_text = format!("{}\0", get_string("MENU_SELECTALL"));
                let selectall_text_utf16: Vec<u16> = selectall_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu_edit,
                    MENU_ITEM_FLAGS(0x00000000),
                    ID_EDIT_SELECTALL as usize,
                    PCWSTR(selectall_text_utf16.as_ptr()),
                );
                let cut_text = format!("{}\0", get_string("MENU_CUT"));
                let cut_text_utf16: Vec<u16> = cut_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu_edit,
                    MENU_ITEM_FLAGS(0x00000000),
                    ID_EDIT_CUT as usize,
                    PCWSTR(cut_text_utf16.as_ptr()),
                );
                let copy_text = format!("{}\0", get_string("MENU_COPY"));
                let copy_text_utf16: Vec<u16> = copy_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu_edit,
                    MENU_ITEM_FLAGS(0x00000000),
                    ID_EDIT_COPY as usize,
                    PCWSTR(copy_text_utf16.as_ptr()),
                );
                let paste_text = format!("{}\0", get_string("MENU_PASTE"));
                let paste_text_utf16: Vec<u16> = paste_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu_edit,
                    MENU_ITEM_FLAGS(0x00000000),
                    ID_EDIT_PASTE as usize,
                    PCWSTR(paste_text_utf16.as_ptr()),
                );
                let edit_text = format!("{}\0", get_string("MENU_EDIT"));
                let edit_text_utf16: Vec<u16> = edit_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu,
                    MENU_ITEM_FLAGS(0x00000010),
                    hmenu_edit.0 as usize,
                    PCWSTR(edit_text_utf16.as_ptr()),
                );

                // Create View menu
                let hmenu_view = CreateMenu().unwrap_or_default();
                let wordwrap_text = format!("{}\0", get_string("MENU_WORDWRAP"));
                let wordwrap_text_utf16: Vec<u16> = wordwrap_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu_view,
                    MENU_ITEM_FLAGS(0x00000000),
                    ID_VIEW_WORDWRAP as usize,
                    PCWSTR(wordwrap_text_utf16.as_ptr()),
                );
                let statusbar_text = format!("{}\0", get_string("MENU_STATUSBAR"));
                let statusbar_text_utf16: Vec<u16> = statusbar_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu_view,
                    MENU_ITEM_FLAGS(0x00000000),
                    ID_VIEW_STATUSBAR as usize,
                    PCWSTR(statusbar_text_utf16.as_ptr()),
                );
                let view_text = format!("{}\0", get_string("MENU_VIEW"));
                let view_text_utf16: Vec<u16> = view_text.encode_utf16().collect();
                let _ = AppendMenuW(
                    hmenu,
                    MENU_ITEM_FLAGS(0x00000010),
                    hmenu_view.0 as usize,
                    PCWSTR(view_text_utf16.as_ptr()),
                );

                // Set menu
                let _ = SetMenu(hwnd, Some(hmenu));

                // Store menu handle
                if let Ok(mut menu_handle) = MENU_HANDLE.lock() {
                    *menu_handle = Some(hmenu.0 as *mut core::ffi::c_void as isize);
                }

                update_wordwrap_menu_check();
                update_statusbar_menu_check();

                // Create status bar components
                let separator_class = "SeparatorClass\0".encode_utf16().collect::<Vec<_>>();
                let separator_hwnd = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(separator_class.as_ptr()),
                    PCWSTR::null(),
                    WINDOW_STYLE(0x40000000 | 0x10000000), // WS_CHILD | WS_VISIBLE
                    0,
                    0,
                    0,
                    2,
                    Some(hwnd),
                    None,
                    Some(HINSTANCE(hinstance.0)),
                    None,
                )
                .unwrap_or_default();
                SetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(16), separator_hwnd.0 as isize);

                // Create status bar sections
                let status_class = "StatusTextClass\0".encode_utf16().collect::<Vec<_>>();
                const SS_LEFT: u32 = 0x0000;
                const SS_RIGHT: u32 = 0x0002;
                const SS_CENTERIMAGE: u32 = 0x0200;

                let char_hwnd = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(status_class.as_ptr()),
                    PCWSTR::null(),
                    WINDOW_STYLE(0x40000000 | 0x10000000 | SS_RIGHT | SS_CENTERIMAGE),
                    0,
                    0,
                    0,
                    20,
                    Some(hwnd),
                    None,
                    Some(HINSTANCE(hinstance.0)),
                    None,
                )
                .unwrap_or_default();

                let pos_hwnd = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(status_class.as_ptr()),
                    PCWSTR::null(),
                    WINDOW_STYLE(0x40000000 | 0x10000000 | SS_LEFT | SS_CENTERIMAGE),
                    0,
                    0,
                    0,
                    20,
                    Some(hwnd),
                    None,
                    Some(HINSTANCE(hinstance.0)),
                    None,
                )
                .unwrap_or_default();

                let encoding_hwnd = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(status_class.as_ptr()),
                    PCWSTR::null(),
                    WINDOW_STYLE(0x40000000 | 0x10000000 | SS_LEFT | SS_CENTERIMAGE),
                    0,
                    0,
                    0,
                    20,
                    Some(hwnd),
                    None,
                    Some(HINSTANCE(hinstance.0)),
                    None,
                )
                .unwrap_or_default();

                // Set font for status bars
                let font_name = "Segoe UI";
                let hfont = CreateFontW(
                    -12,                      // cHeight
                    0,                        // cWidth
                    0,                        // cEscapement
                    0,                        // cOrientation
                    400,                      // cWeight (FW_NORMAL)
                    0,                        // bItalic
                    0,                        // bUnderline
                    0,                        // bStrikeOut
                    FONT_CHARSET(0),          // iCharSet
                    FONT_OUTPUT_PRECISION(0), // iOutPrecision
                    FONT_CLIP_PRECISION(0),   // iClipPrecision
                    FONT_QUALITY(0),          // iQuality
                    0,                        // iPitchAndFamily
                    windows::core::PCWSTR(
                        font_name
                            .encode_utf16()
                            .chain(Some(0))
                            .collect::<Vec<_>>()
                            .as_ptr(),
                    ),
                );
                SendMessageW(
                    char_hwnd,
                    WM_SETFONT,
                    Some(WPARAM(hfont.0 as usize)),
                    Some(LPARAM(1)),
                );
                SendMessageW(
                    pos_hwnd,
                    WM_SETFONT,
                    Some(WPARAM(hfont.0 as usize)),
                    Some(LPARAM(1)),
                );
                SendMessageW(
                    encoding_hwnd,
                    WM_SETFONT,
                    Some(WPARAM(hfont.0 as usize)),
                    Some(LPARAM(1)),
                );

                // Set UTF-8 text
                let encoding_text = "UTF-8\0".encode_utf16().collect::<Vec<_>>();
                let _ = SetWindowTextW(encoding_hwnd, PCWSTR(encoding_text.as_ptr()));

                // Create vertical separators
                let sep1_hwnd = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(separator_class.as_ptr()),
                    PCWSTR::null(),
                    WINDOW_STYLE(0x40000000 | 0x10000000),
                    0,
                    0,
                    0,
                    20,
                    Some(hwnd),
                    None,
                    Some(HINSTANCE(hinstance.0)),
                    None,
                )
                .unwrap_or_default();

                let sep2_hwnd = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(separator_class.as_ptr()),
                    PCWSTR::null(),
                    WINDOW_STYLE(0x40000000 | 0x10000000),
                    0,
                    0,
                    0,
                    20,
                    Some(hwnd),
                    None,
                    Some(HINSTANCE(hinstance.0)),
                    None,
                )
                .unwrap_or_default();

                let sep3_hwnd = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(separator_class.as_ptr()),
                    PCWSTR::null(),
                    WINDOW_STYLE(0x40000000 | 0x10000000),
                    0,
                    0,
                    0,
                    20,
                    Some(hwnd),
                    None,
                    Some(HINSTANCE(hinstance.0)),
                    None,
                )
                .unwrap_or_default();

                // Zoom level display
                let zoom_hwnd = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(status_class.as_ptr()),
                    PCWSTR::null(),
                    WINDOW_STYLE(0x40000000 | 0x10000000 | SS_LEFT | SS_CENTERIMAGE),
                    0,
                    0,
                    32,
                    20,
                    Some(hwnd),
                    None,
                    Some(HINSTANCE(hinstance.0)),
                    None,
                )
                .unwrap_or_default();

                SendMessageW(
                    zoom_hwnd,
                    WM_SETFONT,
                    Some(WPARAM(hfont.0 as usize)),
                    Some(LPARAM(1)),
                );

                let zoom_text = "100%\0".encode_utf16().collect::<Vec<_>>();
                let _ = SetWindowTextW(zoom_hwnd, PCWSTR(zoom_text.as_ptr()));

                let sep4_hwnd = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(separator_class.as_ptr()),
                    PCWSTR::null(),
                    WINDOW_STYLE(0x40000000 | 0x10000000),
                    0,
                    0,
                    0,
                    20,
                    Some(hwnd),
                    None,
                    Some(HINSTANCE(hinstance.0)),
                    None,
                )
                .unwrap_or_default();

                // Line break format display
                let linebreak_hwnd = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(status_class.as_ptr()),
                    PCWSTR::null(),
                    WINDOW_STYLE(0x40000000 | 0x10000000 | SS_LEFT | SS_CENTERIMAGE),
                    0,
                    0,
                    102,
                    20,
                    Some(hwnd),
                    None,
                    Some(HINSTANCE(hinstance.0)),
                    None,
                )
                .unwrap_or_default();

                SendMessageW(
                    linebreak_hwnd,
                    WM_SETFONT,
                    Some(WPARAM(hfont.0 as usize)),
                    Some(LPARAM(1)),
                );

                let linebreak_text = "Windows (CRLF)\0".encode_utf16().collect::<Vec<_>>();
                let _ = SetWindowTextW(linebreak_hwnd, PCWSTR(linebreak_text.as_ptr()));

                // Store handles
                SetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(8), char_hwnd.0 as isize);
                SetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(24), sep1_hwnd.0 as isize);
                SetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(32), pos_hwnd.0 as isize);
                SetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(40), sep2_hwnd.0 as isize);
                SetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(48), encoding_hwnd.0 as isize);
                SetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(56), sep3_hwnd.0 as isize);
                SetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(64), zoom_hwnd.0 as isize);
                SetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(72), sep4_hwnd.0 as isize);
                SetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(80), linebreak_hwnd.0 as isize);

                // Set window icon
                if let Ok(hicon) = LoadIconW(Some(HINSTANCE(hinstance.0)), PCWSTR(1 as *const u16))
                {
                    SendMessageW(
                        hwnd,
                        WM_SETICON,
                        Some(WPARAM(ICON_BIG)),
                        Some(LPARAM(hicon.0 as isize)),
                    );
                    SendMessageW(
                        hwnd,
                        WM_SETICON,
                        Some(WPARAM(ICON_SMALL)),
                        Some(LPARAM(hicon.0 as isize)),
                    );
                }

                // Reset modified flag
                SendMessageW(edit_hwnd, EM_SETMODIFY, Some(WPARAM(0)), Some(LPARAM(0)));

                LRESULT(0)
            }
            WM_GETMINMAXINFO => {
                #[repr(C)]
                #[allow(non_snake_case)]
                struct MINMAXINFO {
                    ptReserved: POINT,
                    ptMaxSize: POINT,
                    ptMaxPosition: POINT,
                    ptMinTrackSize: POINT,
                    ptMaxTrackSize: POINT,
                }

                #[repr(C)]
                struct POINT {
                    x: i32,
                    y: i32,
                }

                let mmi = lparam.0 as *mut MINMAXINFO;
                if !mmi.is_null() {
                    (*mmi).ptMinTrackSize.x = 230;
                }
                LRESULT(0)
            }
            WM_SIZE => {
                let edit_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(0)) as _);

                if edit_hwnd != HWND::default() {
                    let mut rect = RECT::default();
                    let _ = GetClientRect(hwnd, &mut rect);

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

                    let _ = SetWindowPos(
                        edit_hwnd,
                        None,
                        rect.left,
                        rect.top,
                        rect.right - rect.left,
                        edit_height,
                        SET_WINDOW_POS_FLAGS(0x0004), // SWP_NOZORDER
                    );

                    let separator_hwnd =
                        HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(16)) as _);
                    if separator_hwnd != HWND::default() {
                        let _ = SetWindowPos(
                            separator_hwnd,
                            None,
                            rect.left,
                            rect.top + edit_height,
                            rect.right - rect.left,
                            separator_height,
                            SET_WINDOW_POS_FLAGS(0x0004),
                        );
                    }

                    // Position status bar sections
                    let char_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(8)) as _);
                    let sep1_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(24)) as _);
                    let pos_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(32)) as _);
                    let sep2_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(40)) as _);
                    let encoding_hwnd =
                        HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(48)) as _);
                    let sep3_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(56)) as _);
                    let zoom_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(64)) as _);
                    let sep4_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(72)) as _);
                    let linebreak_hwnd =
                        HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(80)) as _);
                    let scrollbar_width = 16;
                    let status_y = rect.top + edit_height + separator_height;

                    let char_width = 80;
                    let separator_width = 2;
                    let pos_width = 122;
                    let zoom_width = 32;
                    let linebreak_width = 102;
                    let encoding_width = 87;
                    let margin = 8;
                    let sep_margin = 8;

                    let total_status_width = char_width
                        + (separator_width + sep_margin * 2)
                        + pos_width
                        + (separator_width + sep_margin * 2)
                        + zoom_width
                        + (separator_width + sep_margin * 2)
                        + linebreak_width
                        + (separator_width + sep_margin * 2)
                        + encoding_width;
                    let start_x = rect.right - scrollbar_width - total_status_width - margin;

                    if char_hwnd != HWND::default() {
                        let _ = SetWindowPos(
                            char_hwnd,
                            None,
                            start_x,
                            status_y,
                            char_width,
                            status_height,
                            SET_WINDOW_POS_FLAGS(0x0004),
                        );
                    }

                    if sep1_hwnd != HWND::default() {
                        let _ = SetWindowPos(
                            sep1_hwnd,
                            None,
                            start_x + char_width + sep_margin,
                            status_y,
                            separator_width,
                            status_height,
                            SET_WINDOW_POS_FLAGS(0x0004),
                        );
                    }

                    if pos_hwnd != HWND::default() {
                        let _ = SetWindowPos(
                            pos_hwnd,
                            None,
                            start_x + char_width + sep_margin + separator_width + sep_margin,
                            status_y,
                            pos_width,
                            status_height,
                            SET_WINDOW_POS_FLAGS(0x0004),
                        );
                    }

                    if sep2_hwnd != HWND::default() {
                        let _ = SetWindowPos(
                            sep2_hwnd,
                            None,
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
                            SET_WINDOW_POS_FLAGS(0x0004),
                        );
                    }

                    if zoom_hwnd != HWND::default() {
                        let _ = SetWindowPos(
                            zoom_hwnd,
                            None,
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
                            zoom_width,
                            status_height,
                            SET_WINDOW_POS_FLAGS(0x0004),
                        );
                    }

                    if sep3_hwnd != HWND::default() {
                        let _ = SetWindowPos(
                            sep3_hwnd,
                            None,
                            start_x
                                + char_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + pos_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + zoom_width
                                + sep_margin,
                            status_y,
                            separator_width,
                            status_height,
                            SET_WINDOW_POS_FLAGS(0x0004),
                        );
                    }

                    if linebreak_hwnd != HWND::default() {
                        let _ = SetWindowPos(
                            linebreak_hwnd,
                            None,
                            start_x
                                + char_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + pos_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + zoom_width
                                + sep_margin
                                + separator_width
                                + sep_margin,
                            status_y,
                            linebreak_width,
                            status_height,
                            SET_WINDOW_POS_FLAGS(0x0004),
                        );
                    }

                    if sep4_hwnd != HWND::default() {
                        let _ = SetWindowPos(
                            sep4_hwnd,
                            None,
                            start_x
                                + char_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + pos_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + zoom_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + linebreak_width
                                + sep_margin,
                            status_y,
                            separator_width,
                            status_height,
                            SET_WINDOW_POS_FLAGS(0x0004),
                        );
                    }

                    if encoding_hwnd != HWND::default() {
                        let _ = SetWindowPos(
                            encoding_hwnd,
                            None,
                            start_x
                                + char_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + pos_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + zoom_width
                                + sep_margin
                                + separator_width
                                + sep_margin
                                + linebreak_width
                                + sep_margin
                                + separator_width
                                + sep_margin,
                            status_y,
                            encoding_width,
                            status_height,
                            SET_WINDOW_POS_FLAGS(0x0004),
                        );
                    }
                }

                LRESULT(0)
            }
            WM_NOTIFY => {
                let edit_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(0)) as _);
                remove_ole_objects(edit_hwnd);
                LRESULT(0)
            }
            WM_PASTE => {
                let edit_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(0)) as _);

                if OpenClipboard(Some(hwnd)).is_ok() {
                    let hdata = GetClipboardData(13); // CF_UNICODETEXT
                    if let Ok(data) = hdata {
                        if !data.0.is_null() {
                            let text_ptr = data.0 as *const u16;
                            let mut len = 0;

                            while *text_ptr.add(len) != 0 {
                                len += 1;
                            }

                            if len > 0 {
                                let text_slice = std::slice::from_raw_parts(text_ptr, len);
                                if let Ok(text) = String::from_utf16(text_slice) {
                                    let text_utf16: Vec<u16> = text.encode_utf16().collect();
                                    SendMessageW(
                                        edit_hwnd,
                                        0x00C2,
                                        Some(WPARAM(1)),
                                        Some(LPARAM(text_utf16.as_ptr() as isize)),
                                    ); // EM_REPLACESEL
                                }
                            }
                        }
                    }

                    let _ = CloseClipboard();
                }

                LRESULT(0)
            }
            WM_SETCURSOR => {
                let hit_test = (lparam.0 & 0xFFFF) as u32;
                const HTCLIENT: u32 = 1;

                if hit_test == HTCLIENT {
                    let cursor_hwnd = HWND(wparam.0 as _);

                    let is_statusbar_visible = if let Ok(visible) = STATUSBAR_VISIBLE.lock() {
                        *visible
                    } else {
                        true
                    };

                    if is_statusbar_visible {
                        #[repr(C)]
                        struct POINT {
                            x: i32,
                            y: i32,
                        }

                        let mut cursor_pt = POINT { x: 0, y: 0 };
                        let _ = GetCursorPos(&mut cursor_pt as *mut _ as *mut _);

                        let mut window_rect = RECT::default();
                        let _ = GetWindowRect(hwnd, &mut window_rect);

                        let mut client_rect = RECT::default();
                        let _ = GetClientRect(hwnd, &mut client_rect);

                        let status_height = 24;
                        let separator_height = 1;
                        let status_total_height = status_height + separator_height;

                        let status_bar_top = window_rect.bottom - status_total_height;

                        if cursor_pt.y >= status_bar_top {
                            if let Ok(cursor) = LoadCursorW(None, IDC_ARROW) {
                                SetCursor(Some(cursor));
                            }
                            return LRESULT(1);
                        }
                    }

                    if cursor_hwnd != hwnd {
                        let char_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(8)) as _);
                        let sep1_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(24)) as _);
                        let pos_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(32)) as _);
                        let sep2_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(40)) as _);
                        let encoding_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(48)) as _);
                        let sep3_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(56)) as _);
                        let zoom_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(64)) as _);
                        let sep4_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(72)) as _);
                        let linebreak_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(80)) as _);
                        let separator_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(16)) as _);

                        if cursor_hwnd == char_hwnd
                            || cursor_hwnd == sep1_hwnd
                            || cursor_hwnd == pos_hwnd
                            || cursor_hwnd == sep2_hwnd
                            || cursor_hwnd == encoding_hwnd
                            || cursor_hwnd == sep3_hwnd
                            || cursor_hwnd == zoom_hwnd
                            || cursor_hwnd == sep4_hwnd
                            || cursor_hwnd == linebreak_hwnd
                            || cursor_hwnd == separator_hwnd
                        {
                            if let Ok(cursor) = LoadCursorW(None, IDC_ARROW) {
                                SetCursor(Some(cursor));
                            }
                            return LRESULT(1);
                        }
                    }
                }

                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CONTEXTMENU => {
                #[repr(C)]
                struct POINT {
                    x: i32,
                    y: i32,
                }

                let mut pt = POINT { x: 0, y: 0 };
                let _ = GetCursorPos(&mut pt as *mut _ as *mut _);
                show_context_menu(hwnd, pt.x, pt.y);
                LRESULT(0)
            }
            WM_COMMAND => {
                let edit_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(0)) as _);
                let cmd_id = wparam.0 as i32;

                match cmd_id {
                    ID_FILE_NEW => {
                        let empty = "\0".encode_utf16().collect::<Vec<_>>();
                        SendMessageW(
                            edit_hwnd,
                            0x000C,
                            Some(WPARAM(0)),
                            Some(LPARAM(empty.as_ptr() as isize)),
                        );

                        let default_filename = get_string("FILE_UNTITLED");
                        let untitled_path = PathBuf::from(default_filename);
                        if let Ok(mut current_file) = CURRENT_FILE.lock() {
                            *current_file = Some(untitled_path);
                        }

                        SendMessageW(edit_hwnd, EM_SETMODIFY, Some(WPARAM(0)), Some(LPARAM(0)));

                        LRESULT(0)
                    }
                    ID_FILE_OPEN => {
                        if let Some((path, selected_encoding)) = file_io::open_file_dialog() {
                            if let Ok((content, detected_encoding)) =
                                file_io::load_file(&path, selected_encoding)
                            {
                                let utf16: Vec<u16> =
                                    content.encode_utf16().chain(std::iter::once(0)).collect();

                                SendMessageW(
                                    edit_hwnd,
                                    0x000C,
                                    Some(WPARAM(0)),
                                    Some(LPARAM(utf16.as_ptr() as isize)),
                                );

                                if let Ok(mut current_file) = CURRENT_FILE.lock() {
                                    *current_file = Some(path.clone());
                                }

                                if let Ok(mut current_encoding) = CURRENT_ENCODING.lock() {
                                    *current_encoding = detected_encoding;
                                }

                                if let Ok(mut saved) = SAVED_CONTENT.lock() {
                                    *saved = content;
                                }

                                SendMessageW(
                                    edit_hwnd,
                                    EM_SETMODIFY,
                                    Some(WPARAM(0)),
                                    Some(LPARAM(0)),
                                );

                                if let Some(filename) = path.file_name() {
                                    if let Some(filename_str) = filename.to_str() {
                                        let app_name = get_string("WINDOW_TITLE");
                                        let title = format!("{} - {}\0", filename_str, app_name);
                                        let title_utf16: Vec<u16> = title.encode_utf16().collect();
                                        let _ = SetWindowTextW(hwnd, PCWSTR(title_utf16.as_ptr()));
                                    }
                                }

                                let char_hwnd =
                                    HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(8)) as _);
                                let pos_hwnd =
                                    HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(32)) as _);
                                let encoding_hwnd =
                                    HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(48)) as _);
                                let zoom_hwnd =
                                    HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(64)) as _);
                                let current_encoding = if let Ok(enc) = CURRENT_ENCODING.lock() {
                                    *enc
                                } else {
                                    FileEncoding::Utf8
                                };
                                update_status_bar(
                                    edit_hwnd,
                                    char_hwnd,
                                    pos_hwnd,
                                    encoding_hwnd,
                                    zoom_hwnd,
                                    current_encoding,
                                );
                            }
                        }
                        LRESULT(0)
                    }
                    ID_FILE_SAVE => {
                        if let Ok(current_file) = CURRENT_FILE.lock() {
                            if let Some(path) = current_file.as_ref() {
                                if is_untitled_file(path) {
                                    let path_clone = path.clone();
                                    drop(current_file);
                                    let current_encoding = if let Ok(enc) = CURRENT_ENCODING.lock()
                                    {
                                        *enc
                                    } else {
                                        FileEncoding::Utf8
                                    };
                                    if let Some((new_path, encoding)) = file_io::save_file_dialog(
                                        current_encoding,
                                        Some(&path_clone),
                                    ) {
                                        let text_len = SendMessageW(
                                            edit_hwnd,
                                            0x000E,
                                            Some(WPARAM(0)),
                                            Some(LPARAM(0)),
                                        )
                                        .0
                                            as usize;
                                        let text = if text_len > 0 {
                                            let mut buffer: Vec<u16> = vec![0; text_len + 1];
                                            SendMessageW(
                                                edit_hwnd,
                                                0x000D,
                                                Some(WPARAM((text_len + 1) as usize)),
                                                Some(LPARAM(buffer.as_mut_ptr() as isize)),
                                            );
                                            String::from_utf16(&buffer[..text_len])
                                                .unwrap_or_default()
                                        } else {
                                            String::new()
                                        };

                                        let _ = file_io::save_file(&new_path, &text, encoding);
                                        *CURRENT_FILE.lock().unwrap() = Some(new_path.clone());
                                        if let Ok(mut current_encoding) = CURRENT_ENCODING.lock() {
                                            *current_encoding = encoding;
                                        }
                                        if let Ok(mut saved) = SAVED_CONTENT.lock() {
                                            *saved = text;
                                        }
                                        SendMessageW(
                                            edit_hwnd,
                                            EM_SETMODIFY,
                                            Some(WPARAM(0)),
                                            Some(LPARAM(0)),
                                        );
                                        if let Some(filename) = new_path.file_name() {
                                            if let Some(filename_str) = filename.to_str() {
                                                let app_name = get_string("WINDOW_TITLE");
                                                let title =
                                                    format!("{} - {}\0", filename_str, app_name);
                                                let title_utf16: Vec<u16> =
                                                    title.encode_utf16().collect();
                                                let _ = SetWindowTextW(
                                                    hwnd,
                                                    PCWSTR(title_utf16.as_ptr()),
                                                );
                                            }
                                        }
                                        let char_hwnd =
                                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(8))
                                                as _);
                                        let pos_hwnd = HWND(GetWindowLongPtrW(
                                            hwnd,
                                            WINDOW_LONG_PTR_INDEX(32),
                                        )
                                            as _);
                                        let encoding_hwnd = HWND(GetWindowLongPtrW(
                                            hwnd,
                                            WINDOW_LONG_PTR_INDEX(48),
                                        )
                                            as _);
                                        let zoom_hwnd = HWND(GetWindowLongPtrW(
                                            hwnd,
                                            WINDOW_LONG_PTR_INDEX(64),
                                        )
                                            as _);
                                        let current_encoding =
                                            if let Ok(enc) = CURRENT_ENCODING.lock() {
                                                *enc
                                            } else {
                                                FileEncoding::Utf8
                                            };
                                        update_status_bar(
                                            edit_hwnd,
                                            char_hwnd,
                                            pos_hwnd,
                                            encoding_hwnd,
                                            zoom_hwnd,
                                            current_encoding,
                                        );
                                    }
                                } else {
                                    let encoding = if let Ok(enc) = CURRENT_ENCODING.lock() {
                                        *enc
                                    } else {
                                        FileEncoding::Utf8
                                    };

                                    let text_len = SendMessageW(
                                        edit_hwnd,
                                        0x000E,
                                        Some(WPARAM(0)),
                                        Some(LPARAM(0)),
                                    )
                                    .0 as usize;
                                    let text = if text_len > 0 {
                                        let mut buffer: Vec<u16> = vec![0; text_len + 1];
                                        SendMessageW(
                                            edit_hwnd,
                                            0x000D,
                                            Some(WPARAM((text_len + 1) as usize)),
                                            Some(LPARAM(buffer.as_mut_ptr() as isize)),
                                        );
                                        String::from_utf16(&buffer[..text_len]).unwrap_or_default()
                                    } else {
                                        String::new()
                                    };

                                    let _ = file_io::save_file(path, &text, encoding);
                                    if let Ok(mut saved) = SAVED_CONTENT.lock() {
                                        *saved = text;
                                    }
                                    SendMessageW(
                                        edit_hwnd,
                                        EM_SETMODIFY,
                                        Some(WPARAM(0)),
                                        Some(LPARAM(0)),
                                    );
                                    if let Some(filename) = path.file_name() {
                                        if let Some(filename_str) = filename.to_str() {
                                            let app_name = get_string("WINDOW_TITLE");
                                            let title =
                                                format!("{} - {}\0", filename_str, app_name);
                                            let title_utf16: Vec<u16> =
                                                title.encode_utf16().collect();
                                            let _ =
                                                SetWindowTextW(hwnd, PCWSTR(title_utf16.as_ptr()));
                                        }
                                    }
                                }
                            }
                        }
                        LRESULT(0)
                    }
                    ID_FILE_SAVEAS => {
                        let current_file_path = if let Ok(file) = CURRENT_FILE.lock() {
                            file.clone()
                        } else {
                            None
                        };
                        let current_encoding = if let Ok(enc) = CURRENT_ENCODING.lock() {
                            *enc
                        } else {
                            FileEncoding::Utf8
                        };
                        if let Some((new_path, encoding)) =
                            file_io::save_file_dialog(current_encoding, current_file_path.as_ref())
                        {
                            let text_len =
                                SendMessageW(edit_hwnd, 0x000E, Some(WPARAM(0)), Some(LPARAM(0))).0
                                    as usize;
                            let text = if text_len > 0 {
                                let mut buffer: Vec<u16> = vec![0; text_len + 1];
                                SendMessageW(
                                    edit_hwnd,
                                    0x000D,
                                    Some(WPARAM((text_len + 1) as usize)),
                                    Some(LPARAM(buffer.as_mut_ptr() as isize)),
                                );
                                String::from_utf16(&buffer[..text_len]).unwrap_or_default()
                            } else {
                                String::new()
                            };

                            let _ = file_io::save_file(&new_path, &text, encoding);
                            *CURRENT_FILE.lock().unwrap() = Some(new_path.clone());
                            if let Ok(mut current_encoding) = CURRENT_ENCODING.lock() {
                                *current_encoding = encoding;
                            }
                            if let Ok(mut saved) = SAVED_CONTENT.lock() {
                                *saved = text.clone();
                            }
                            SendMessageW(edit_hwnd, EM_SETMODIFY, Some(WPARAM(0)), Some(LPARAM(0)));
                            if let Some(filename) = new_path.file_name() {
                                if let Some(filename_str) = filename.to_str() {
                                    let app_name = get_string("WINDOW_TITLE");
                                    let title = format!("{} - {}\0", filename_str, app_name);
                                    let title_utf16: Vec<u16> = title.encode_utf16().collect();
                                    let _ = SetWindowTextW(hwnd, PCWSTR(title_utf16.as_ptr()));
                                }
                            }
                            let char_hwnd =
                                HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(8)) as _);
                            let pos_hwnd =
                                HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(32)) as _);
                            let encoding_hwnd =
                                HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(48)) as _);
                            let zoom_hwnd =
                                HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(64)) as _);
                            update_status_bar(
                                edit_hwnd,
                                char_hwnd,
                                pos_hwnd,
                                encoding_hwnd,
                                zoom_hwnd,
                                encoding,
                            );
                        }
                        LRESULT(0)
                    }
                    ID_FILE_EXIT => {
                        PostQuitMessage(0);
                        LRESULT(0)
                    }
                    ID_EDIT_SELECTALL => {
                        SendMessageW(
                            edit_hwnd,
                            EM_SETSEL as u32,
                            Some(WPARAM(0)),
                            Some(LPARAM(-1)),
                        );
                        LRESULT(0)
                    }
                    ID_EDIT_CUT => {
                        SendMessageW(edit_hwnd, WM_CUT, Some(WPARAM(0)), Some(LPARAM(0)));
                        LRESULT(0)
                    }
                    ID_EDIT_COPY => {
                        SendMessageW(edit_hwnd, WM_COPY, Some(WPARAM(0)), Some(LPARAM(0)));
                        LRESULT(0)
                    }
                    ID_EDIT_PASTE => {
                        SendMessageW(hwnd, WM_PASTE, Some(WPARAM(0)), Some(LPARAM(0)));
                        LRESULT(0)
                    }
                    ID_VIEW_WORDWRAP => {
                        toggle_word_wrap(edit_hwnd);
                        LRESULT(0)
                    }
                    ID_VIEW_STATUSBAR => {
                        let new_visibility = if let Ok(mut visible) = STATUSBAR_VISIBLE.lock() {
                            *visible = !*visible;
                            *visible
                        } else {
                            true
                        };

                        let char_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(8)) as _);
                        let sep1_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(24)) as _);
                        let pos_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(32)) as _);
                        let sep2_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(40)) as _);
                        let encoding_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(48)) as _);
                        let sep3_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(56)) as _);
                        let zoom_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(64)) as _);
                        let sep4_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(72)) as _);
                        let linebreak_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(80)) as _);
                        let separator_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(16)) as _);

                        let show_cmd = if new_visibility {
                            SHOW_WINDOW_CMD(5) // SW_SHOW
                        } else {
                            SHOW_WINDOW_CMD(0) // SW_HIDE
                        };
                        let _ = ShowWindow(char_hwnd, show_cmd);
                        let _ = ShowWindow(sep1_hwnd, show_cmd);
                        let _ = ShowWindow(pos_hwnd, show_cmd);
                        let _ = ShowWindow(sep2_hwnd, show_cmd);
                        let _ = ShowWindow(encoding_hwnd, show_cmd);
                        let _ = ShowWindow(sep3_hwnd, show_cmd);
                        let _ = ShowWindow(zoom_hwnd, show_cmd);
                        let _ = ShowWindow(sep4_hwnd, show_cmd);
                        let _ = ShowWindow(linebreak_hwnd, show_cmd);
                        let _ = ShowWindow(separator_hwnd, show_cmd);

                        if let Ok(menu_handle) = MENU_HANDLE.lock() {
                            if let Some(hmenu_isize) = *menu_handle {
                                let check_state = if new_visibility {
                                    MENU_ITEM_FLAGS(0x00000008)
                                } else {
                                    MENU_ITEM_FLAGS(0x00000000)
                                };
                                let _ = CheckMenuItem(
                                    HMENU(hmenu_isize as *mut core::ffi::c_void),
                                    ID_VIEW_STATUSBAR as u32,
                                    check_state.0,
                                );
                            }
                        }

                        let mut rect = RECT::default();
                        let _ = GetClientRect(hwnd, &mut rect);
                        let width = rect.right - rect.left;
                        let height = rect.bottom - rect.top;
                        SendMessageW(
                            hwnd,
                            WM_SIZE,
                            Some(WPARAM(0)),
                            Some(LPARAM(((height as isize) << 16) | (width as isize))),
                        );

                        let _ = InvalidateRect(Some(hwnd), None, true);

                        LRESULT(0)
                    }
                    _ => DefWindowProcW(hwnd, msg, wparam, lparam),
                }
            }
            0x0101 | 0x0202 => {
                // WM_KEYUP | WM_LBUTTONUP
                if let Ok(visible) = STATUSBAR_VISIBLE.lock() {
                    if *visible {
                        let edit_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(0)) as _);
                        let char_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(8)) as _);
                        let pos_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(32)) as _);
                        let encoding_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(48)) as _);
                        let zoom_hwnd =
                            HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(64)) as _);
                        let current_encoding = if let Ok(enc) = CURRENT_ENCODING.lock() {
                            *enc
                        } else {
                            FileEncoding::Utf8
                        };
                        update_status_bar(
                            edit_hwnd,
                            char_hwnd,
                            pos_hwnd,
                            encoding_hwnd,
                            zoom_hwnd,
                            current_encoding,
                        );
                    }
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
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
}

fn main() {
    init_language();

    let default_filename = get_string("FILE_UNTITLED");
    let untitled_path = PathBuf::from(default_filename);
    *CURRENT_FILE.lock().unwrap() = Some(untitled_path);

    unsafe {
        let hinstance = GetModuleHandleW(None).unwrap_or_default();
        let class_name = "NotepadWindowClass\0".encode_utf16().collect::<Vec<_>>();

        const COLOR_BTNFACE: SYS_COLOR_INDEX = SYS_COLOR_INDEX(15);

        let hicon =
            LoadIconW(Some(HINSTANCE(hinstance.0)), PCWSTR(std::ptr::null())).unwrap_or_default();

        let wnd_class = WNDCLASSW {
            style: WNDCLASS_STYLES(0x0001 | 0x0002), // CS_VREDRAW | CS_HREDRAW
            lpfnWndProc: Some(window_proc),
            cbClsExtra: 0,
            cbWndExtra: (std::mem::size_of::<isize>() * 11) as i32,
            hInstance: HINSTANCE(hinstance.0),
            hIcon: hicon,
            hCursor: Default::default(),
            hbrBackground: HBRUSH(GetSysColorBrush(COLOR_BTNFACE).0),
            lpszMenuName: PCWSTR::null(),
            lpszClassName: PCWSTR(class_name.as_ptr()),
        };

        let _ = RegisterClassW(&wnd_class);

        status_bar::register_status_bar_classes();

        let window_title_str = format!("{}\0", get_string("WINDOW_TITLE"));
        let window_title = window_title_str.encode_utf16().collect::<Vec<_>>();

        const WS_THICKFRAME: u32 = 0x00040000;
        const CW_USEDEFAULT: i32 = i32::MIN;
        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE(0),
            PCWSTR(class_name.as_ptr()),
            PCWSTR(window_title.as_ptr()),
            WINDOW_STYLE(0x00CF0000 | 0x10000000 | WS_THICKFRAME), // WS_OVERLAPPEDWINDOW | WS_VISIBLE
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            800,
            600,
            None,
            None,
            Some(HINSTANCE(hinstance.0)),
            None,
        )
        .unwrap_or_default();

        let app_name = get_string("WINDOW_TITLE");
        let default_filename = get_string("FILE_UNTITLED");
        let initial_title = format!("{} - {}\0", default_filename, app_name);
        let initial_title_utf16: Vec<u16> = initial_title.encode_utf16().collect();
        let _ = SetWindowTextW(hwnd, PCWSTR(initial_title_utf16.as_ptr()));

        let mut msg = MSG::default();
        while GetMessageW(&mut msg, None, 0, 0).as_bool() {
            if msg.message == WM_KEYDOWN
                && (msg.wParam.0 as i32 == 0x45 || msg.wParam.0 as i32 == 0x52)
            {
                let ctrl_pressed = (GetKeyState(0x11) as u16 & 0x8000) != 0;
                if ctrl_pressed {
                    continue;
                }
            }

            if msg.message == WM_KEYDOWN && msg.wParam.0 as i32 == 0x53 {
                let ctrl_pressed = (GetKeyState(0x11) as u16 & 0x8000) != 0;
                if ctrl_pressed {
                    let edit_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(0)) as _);
                    if let Ok(current_file) = CURRENT_FILE.lock() {
                        if let Some(path) = current_file.as_ref() {
                            if is_untitled_file(path) {
                                let path_clone = path.clone();
                                drop(current_file);
                                let current_encoding = if let Ok(enc) = CURRENT_ENCODING.lock() {
                                    *enc
                                } else {
                                    FileEncoding::Utf8
                                };
                                if let Some((new_path, encoding)) =
                                    file_io::save_file_dialog(current_encoding, Some(&path_clone))
                                {
                                    let text_len = SendMessageW(
                                        edit_hwnd,
                                        0x000E,
                                        Some(WPARAM(0)),
                                        Some(LPARAM(0)),
                                    )
                                    .0 as usize;
                                    let text = if text_len > 0 {
                                        let mut buffer: Vec<u16> = vec![0; text_len + 1];
                                        SendMessageW(
                                            edit_hwnd,
                                            0x000D,
                                            Some(WPARAM((text_len + 1) as usize)),
                                            Some(LPARAM(buffer.as_mut_ptr() as isize)),
                                        );
                                        String::from_utf16(&buffer[..text_len]).unwrap_or_default()
                                    } else {
                                        String::new()
                                    };

                                    let _ = file_io::save_file(&new_path, &text, encoding);
                                    *CURRENT_FILE.lock().unwrap() = Some(new_path.clone());
                                    if let Ok(mut current_encoding) = CURRENT_ENCODING.lock() {
                                        *current_encoding = encoding;
                                    }
                                    if let Ok(mut saved) = SAVED_CONTENT.lock() {
                                        *saved = text.clone();
                                    }
                                    SendMessageW(
                                        edit_hwnd,
                                        EM_SETMODIFY,
                                        Some(WPARAM(0)),
                                        Some(LPARAM(0)),
                                    );
                                    if let Some(filename) = new_path.file_name() {
                                        if let Some(filename_str) = filename.to_str() {
                                            let app_name = get_string("WINDOW_TITLE");
                                            let title =
                                                format!("{} - {}\0", filename_str, app_name);
                                            let title_utf16: Vec<u16> =
                                                title.encode_utf16().collect();
                                            let _ =
                                                SetWindowTextW(hwnd, PCWSTR(title_utf16.as_ptr()));
                                        }
                                    }
                                    // Update status bar
                                    let char_hwnd = HWND(GetWindowLongPtrW(
                                        hwnd,
                                        WINDOW_LONG_PTR_INDEX(8),
                                    )
                                        as _);
                                    let pos_hwnd = HWND(GetWindowLongPtrW(
                                        hwnd,
                                        WINDOW_LONG_PTR_INDEX(32),
                                    ) as _);
                                    let encoding_hwnd =
                                        HWND(
                                            GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(48)) as _
                                        );
                                    let zoom_hwnd =
                                        HWND(
                                            GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(64)) as _
                                        );
                                    update_status_bar(
                                        edit_hwnd,
                                        char_hwnd,
                                        pos_hwnd,
                                        encoding_hwnd,
                                        zoom_hwnd,
                                        encoding,
                                    );
                                }
                            } else {
                                let text_len = SendMessageW(
                                    edit_hwnd,
                                    0x000E,
                                    Some(WPARAM(0)),
                                    Some(LPARAM(0)),
                                )
                                .0 as usize;
                                let text = if text_len > 0 {
                                    let mut buffer: Vec<u16> = vec![0; text_len + 1];
                                    SendMessageW(
                                        edit_hwnd,
                                        0x000D,
                                        Some(WPARAM((text_len + 1) as usize)),
                                        Some(LPARAM(buffer.as_mut_ptr() as isize)),
                                    );
                                    String::from_utf16(&buffer[..text_len]).unwrap_or_default()
                                } else {
                                    String::new()
                                };

                                let current_encoding = if let Ok(enc) = CURRENT_ENCODING.lock() {
                                    *enc
                                } else {
                                    FileEncoding::Utf8
                                };
                                let _ = file_io::save_file(path, &text, current_encoding);
                                if let Ok(mut saved) = SAVED_CONTENT.lock() {
                                    *saved = text.clone();
                                }
                                SendMessageW(
                                    edit_hwnd,
                                    EM_SETMODIFY,
                                    Some(WPARAM(0)),
                                    Some(LPARAM(0)),
                                );
                                if let Some(filename) = path.file_name() {
                                    if let Some(filename_str) = filename.to_str() {
                                        let app_name = get_string("WINDOW_TITLE");
                                        let title = format!("{} - {}\0", filename_str, app_name);
                                        let title_utf16: Vec<u16> = title.encode_utf16().collect();
                                        let _ = SetWindowTextW(hwnd, PCWSTR(title_utf16.as_ptr()));
                                    }
                                }
                            }
                        }
                    }
                    continue;
                }
            }

            if msg.message == WM_KEYDOWN && msg.wParam.0 as i32 == 0x56 {
                let ctrl_pressed = (GetKeyState(0x11) as u16 & 0x8000) != 0;
                if ctrl_pressed {
                    let edit_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(0)) as _);

                    if OpenClipboard(Some(hwnd)).is_ok() {
                        let hdata = GetClipboardData(13);
                        if let Ok(data) = hdata {
                            if !data.0.is_null() {
                                let text_ptr = data.0 as *const u16;
                                let mut len = 0;

                                while *text_ptr.add(len) != 0 {
                                    len += 1;
                                }

                                if len > 0 {
                                    let text_slice = std::slice::from_raw_parts(text_ptr, len);
                                    if let Ok(text) = String::from_utf16(text_slice) {
                                        let text_utf16: Vec<u16> =
                                            text.encode_utf16().chain(std::iter::once(0)).collect();
                                        SendMessageW(
                                            edit_hwnd,
                                            0x00C2,
                                            Some(WPARAM(1)),
                                            Some(LPARAM(text_utf16.as_ptr() as isize)),
                                        );
                                    }
                                }
                            }
                        }

                        let _ = CloseClipboard();
                    }

                    continue;
                }
            }

            let _ = TranslateMessage(&msg);
            let _ = DispatchMessageW(&msg);

            let edit_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(0)) as _);

            if let Ok(visible) = STATUSBAR_VISIBLE.lock() {
                if *visible {
                    let char_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(8)) as _);
                    let pos_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(32)) as _);
                    let encoding_hwnd =
                        HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(48)) as _);
                    let zoom_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(64)) as _);
                    let current_encoding = if let Ok(enc) = CURRENT_ENCODING.lock() {
                        *enc
                    } else {
                        FileEncoding::Utf8
                    };
                    status_bar::update_status_bar(
                        edit_hwnd,
                        char_hwnd,
                        pos_hwnd,
                        encoding_hwnd,
                        zoom_hwnd,
                        current_encoding,
                    );
                }
            }

            update_title_if_needed(hwnd, edit_hwnd);
        }
    }
}
