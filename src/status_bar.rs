use crate::file_io::FileEncoding;
use crate::i18n::get_string;
use crate::line_column::calculate_line_column;
use std::sync::Mutex;
use std::sync::atomic::{AtomicBool, Ordering};
use windows_sys::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows_sys::Win32::Graphics::Gdi::InvalidateRect;
use windows_sys::Win32::Graphics::Gdi::{
    BeginPaint, CreatePen, EndPaint, GetSysColorBrush, LineTo, MoveToEx, SelectObject,
};
use windows_sys::Win32::System::LibraryLoader::GetModuleHandleW;
use windows_sys::Win32::UI::WindowsAndMessaging::{
    CS_HREDRAW, CS_VREDRAW, DefWindowProcW, GetClientRect, IDC_ARROW, LoadCursorW, SendMessageW,
    SetCursor, SetWindowTextW, WM_PAINT, WM_SETCURSOR, WNDCLASSW,
};

pub const EM_GETSEL: u32 = 0x00B0;
pub const EM_LINEFROMCHAR: u32 = 0x00C9;
pub const EM_GETZOOM: u32 = 0x04E0;
pub const WM_GETTEXTLENGTH: u32 = 0x000E;

// Helper function to convert raw SendMessageW result to i32
#[inline]
fn msg_as_i32(result: isize) -> i32 {
    result as i32
}

// Helper function to convert raw SendMessageW result to usize
#[inline]
fn msg_as_usize(result: isize) -> usize {
    result as usize
}

// Count newlines as 1 character (\r\n = 1 char) or 2 characters (\r\n = 2 chars)
static COUNT_NEWLINE_AS_ONE: AtomicBool = AtomicBool::new(true);

// Cache for previous status bar values
static LAST_STATUS: Mutex<Option<(i32, i32, i32, i32)>> = Mutex::new(None);

// Separator window procedure for thin light gray lines (vertical or horizontal)
pub extern "system" fn separator_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        match msg {
            WM_SETCURSOR => {
                // Set cursor to default arrow when hovering over status bar
                let cursor = LoadCursorW(std::ptr::null_mut(), IDC_ARROW);
                SetCursor(cursor);
                1
            }
            WM_PAINT => {
                let mut ps = std::mem::zeroed();
                BeginPaint(hwnd, &mut ps);

                let mut rect: RECT = std::mem::zeroed();
                GetClientRect(hwnd, &mut rect);

                // Gray color (RGB: 210, 209, 208)
                let separator_color = 0x00D0D1D2u32;
                let pen = CreatePen(0, 1, separator_color);
                SelectObject(ps.hdc, pen as *mut std::ffi::c_void);

                let width = rect.right - rect.left;
                let height = rect.bottom - rect.top;

                if height > width {
                    // Vertical line - draw in the middle
                    let x = width / 2;
                    MoveToEx(ps.hdc, x, rect.top, std::ptr::null_mut());
                    LineTo(ps.hdc, x, rect.bottom);
                } else {
                    // Horizontal line - draw in the middle
                    let y = height / 2;
                    MoveToEx(ps.hdc, rect.left, y, std::ptr::null_mut());
                    LineTo(ps.hdc, rect.right, y);
                }

                EndPaint(hwnd, &ps);
                0
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

// Status text window procedure (for status bar text labels)
pub extern "system" fn status_text_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        match msg {
            WM_SETCURSOR => {
                // Set cursor to default arrow
                let cursor = LoadCursorW(std::ptr::null_mut(), IDC_ARROW);
                SetCursor(cursor);
                1
            }
            WM_PAINT => {
                // Paint text using STATIC control behavior
                use windows_sys::Win32::Graphics::Gdi::{
                    BeginPaint, DrawTextW, EndPaint, FillRect, SetBkMode, SetTextColor, TRANSPARENT,
                };
                use windows_sys::Win32::UI::WindowsAndMessaging::GetWindowLongPtrW;

                const DT_LEFT: u32 = 0x00000000;
                const DT_RIGHT: u32 = 0x00000002;
                const DT_SINGLELINE: u32 = 0x00000020;
                const DT_VCENTER: u32 = 0x00000004;
                const GWL_STYLE: i32 = -16;

                let mut ps = std::mem::zeroed();
                BeginPaint(hwnd, &mut ps);

                let mut rect: RECT = std::mem::zeroed();
                GetClientRect(hwnd, &mut rect);

                // Fill background
                const COLOR_BTNFACE: i32 = 15;
                let brush = GetSysColorBrush(COLOR_BTNFACE);
                FillRect(ps.hdc, &rect, brush);

                // Get window text
                const WM_GETTEXT: u32 = 0x000D;
                const WM_GETTEXTLENGTH: u32 = 0x000E;
                let text_len = windows_sys::Win32::UI::WindowsAndMessaging::SendMessageW(
                    hwnd,
                    WM_GETTEXTLENGTH,
                    0,
                    0,
                ) as usize;

                if text_len > 0 {
                    let mut buffer = vec![0u16; text_len + 1];
                    windows_sys::Win32::UI::WindowsAndMessaging::SendMessageW(
                        hwnd,
                        WM_GETTEXT,
                        buffer.len(),
                        buffer.as_mut_ptr() as isize,
                    );

                    // Set text properties
                    SetBkMode(ps.hdc, TRANSPARENT as i32);
                    const COLOR_BTNTEXT: i32 = 18;
                    SetTextColor(
                        ps.hdc,
                        windows_sys::Win32::Graphics::Gdi::GetSysColor(COLOR_BTNTEXT),
                    );

                    // Create and set small font for status bar
                    use windows_sys::Win32::Graphics::Gdi::{CreateFontW, DeleteObject, FW_NORMAL};
                    const DEFAULT_CHARSET: u32 = 1;
                    const OUT_DEFAULT_PRECIS: u32 = 0;
                    const CLIP_DEFAULT_PRECIS: u32 = 0;
                    const DEFAULT_QUALITY: u32 = 0;
                    const DEFAULT_PITCH: u32 = 0;
                    const FF_DONTCARE: u32 = 0;

                    let font_name = "Segoe UI\0".encode_utf16().collect::<Vec<_>>();
                    let font = CreateFontW(
                        -12, // Height (negative = point size)
                        0,   // Width
                        0,   // Escapement
                        0,   // Orientation
                        FW_NORMAL as i32,
                        0, // Italic
                        0, // Underline
                        0, // StrikeOut
                        DEFAULT_CHARSET,
                        OUT_DEFAULT_PRECIS,
                        CLIP_DEFAULT_PRECIS,
                        DEFAULT_QUALITY,
                        (DEFAULT_PITCH | FF_DONTCARE) << 8,
                        font_name.as_ptr(),
                    );
                    let old_font = SelectObject(ps.hdc, font as *mut std::ffi::c_void);

                    // Check style for alignment
                    const SS_RIGHT: isize = 0x0002;
                    let style = GetWindowLongPtrW(hwnd, GWL_STYLE);
                    let format = if (style & SS_RIGHT) != 0 {
                        DT_RIGHT | DT_SINGLELINE | DT_VCENTER
                    } else {
                        DT_LEFT | DT_SINGLELINE | DT_VCENTER
                    };

                    DrawTextW(ps.hdc, buffer.as_ptr(), text_len as i32, &mut rect, format);

                    // Restore old font and delete created font
                    SelectObject(ps.hdc, old_font);
                    DeleteObject(font);
                }

                EndPaint(hwnd, &ps);
                0
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

// Register status bar window classes
pub unsafe fn register_status_bar_classes() {
    unsafe {
        let hinstance = GetModuleHandleW(std::ptr::null());
        const COLOR_BTNFACE: i32 = 15;

        // Register separator window class
        let separator_class_name = "SeparatorClass\0".encode_utf16().collect::<Vec<_>>();
        let separator_class = WNDCLASSW {
            style: CS_VREDRAW | CS_HREDRAW,
            lpfnWndProc: Some(separator_proc),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: hinstance,
            hIcon: std::ptr::null_mut(),
            hCursor: std::ptr::null_mut(),
            hbrBackground: GetSysColorBrush(COLOR_BTNFACE),
            lpszMenuName: std::ptr::null(),
            lpszClassName: separator_class_name.as_ptr(),
        };
        windows_sys::Win32::UI::WindowsAndMessaging::RegisterClassW(&separator_class);

        // Register status text window class
        let status_text_class_name = "StatusTextClass\0".encode_utf16().collect::<Vec<_>>();
        let status_text_class = WNDCLASSW {
            style: CS_VREDRAW | CS_HREDRAW,
            lpfnWndProc: Some(status_text_proc),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: hinstance,
            hIcon: std::ptr::null_mut(),
            hCursor: std::ptr::null_mut(),
            hbrBackground: GetSysColorBrush(COLOR_BTNFACE),
            lpszMenuName: std::ptr::null(),
            lpszClassName: status_text_class_name.as_ptr(),
        };
        windows_sys::Win32::UI::WindowsAndMessaging::RegisterClassW(&status_text_class);
    }
}

pub fn update_status_bar(
    edit_hwnd: HWND,
    char_hwnd: HWND,
    pos_hwnd: HWND,
    encoding_hwnd: HWND,
    zoom_hwnd: HWND,
    current_encoding: FileEncoding,
) {
    unsafe {
        if edit_hwnd != HWND::default() {
            // Get cursor position using EM_GETSEL
            // EM_GETSEL expects wparam=pointer to start, lparam=pointer to end
            let mut start_pos: i32 = 0;
            let mut end_pos: i32 = 0;
            SendMessageW(
                edit_hwnd,
                EM_GETSEL,
                &mut start_pos as *mut i32 as usize,
                &mut end_pos as *mut i32 as isize,
            );

            // Get text by selecting all and using EM_GETSEL with a custom range
            // First, get text length
            let text_length = msg_as_i32(SendMessageW(edit_hwnd, WM_GETTEXTLENGTH, 0, 0));

            let text_str = if text_length > 0 {
                // Allocate buffer and get text using WM_GETTEXT (0x000D)
                let mut buffer = vec![0u16; (text_length + 1) as usize];
                let actual_len = msg_as_usize(SendMessageW(
                    edit_hwnd,
                    0x000D,
                    (text_length + 1) as usize,
                    buffer.as_mut_ptr() as isize,
                ));

                if actual_len > 0 {
                    String::from_utf16_lossy(&buffer[..actual_len.min(text_length as usize)])
                        .to_string()
                } else {
                    String::new()
                }
            } else {
                String::new()
            };

            // Convert char position (Windows treats \r\n as 1 char) to UTF-16 code unit position
            let mut utf16_pos = 0i32;
            let mut char_idx = 0i32;
            let mut chars_iter = text_str.chars().peekable();

            while let Some(ch) = chars_iter.next() {
                if char_idx >= start_pos {
                    break;
                }

                if ch == '\r' && chars_iter.peek() == Some(&'\n') {
                    // \r\n counts as 1 character in Windows
                    chars_iter.next(); // consume the \n
                    utf16_pos += 2; // but occupies 2 UTF-16 code units
                } else {
                    utf16_pos += if ch > '\u{FFFF}' { 2 } else { 1 };
                }

                char_idx += 1;
            }

            // Calculate line and column using the dedicated function
            let (display_line, display_col) = calculate_line_column(&text_str, utf16_pos);

            // Get total character count
            let char_count = if COUNT_NEWLINE_AS_ONE.load(Ordering::SeqCst) {
                // Count characters, treating \r\n as 1 character
                let mut count = 0i32;
                let mut chars = text_str.chars().peekable();
                while let Some(ch) = chars.next() {
                    if ch == '\r' && chars.peek() == Some(&'\n') {
                        chars.next(); // Skip the \n
                        count += 1;
                    } else if ch != '\0' {
                        count += 1;
                    }
                }
                count
            } else {
                // Count \r\n as 2 characters
                text_length
            };

            // Get zoom level
            let mut numerator: i32 = 0;
            let mut denominator: i32 = 0;
            SendMessageW(
                edit_hwnd,
                EM_GETZOOM,
                &mut numerator as *mut i32 as usize,
                &mut denominator as *mut i32 as isize,
            );

            let zoom_percent = if denominator != 0 {
                (numerator * 100) / denominator
            } else {
                100 // Default 100%
            };

            // Check if values changed
            let current_status = (display_line, display_col, char_count, zoom_percent);
            let mut last = LAST_STATUS.lock().unwrap();

            if *last != Some(current_status) {
                *last = Some(current_status);

                // Update character count
                let char_format = get_string("STATUS_CHAR_COUNT");
                let char_text = format!(
                    "{}\0",
                    char_format.replace("{count}", &char_count.to_string())
                );
                let char_utf16: Vec<u16> = char_text.encode_utf16().collect();
                SetWindowTextW(char_hwnd, char_utf16.as_ptr());
                InvalidateRect(char_hwnd, std::ptr::null(), 1);

                // Update line and column in same section
                let pos_format = get_string("STATUS_LINE_COL");
                let pos_text = format!(
                    "{}\0",
                    pos_format
                        .replace("{line}", &display_line.to_string())
                        .replace("{col}", &display_col.to_string())
                );
                let pos_utf16: Vec<u16> = pos_text.encode_utf16().collect();
                SetWindowTextW(pos_hwnd, pos_utf16.as_ptr());
                InvalidateRect(pos_hwnd, std::ptr::null(), 1);

                // Update encoding display
                let encoding_text = match current_encoding {
                    FileEncoding::Utf8 => "UTF-8\0".to_string(),
                    FileEncoding::ShiftJis => "Shift-JIS\0".to_string(),
                    FileEncoding::Auto => format!("{}\0", get_string("ENCODING_AUTO")),
                };
                let encoding_utf16: Vec<u16> = encoding_text.encode_utf16().collect();
                SetWindowTextW(encoding_hwnd, encoding_utf16.as_ptr());
                InvalidateRect(encoding_hwnd, std::ptr::null(), 1);

                // Update zoom level
                let zoom_text = format!("{}%\0", zoom_percent);
                let zoom_utf16: Vec<u16> = zoom_text.encode_utf16().collect();
                SetWindowTextW(zoom_hwnd, zoom_utf16.as_ptr());
                InvalidateRect(zoom_hwnd, std::ptr::null(), 1);
            }
        }
    }
}
