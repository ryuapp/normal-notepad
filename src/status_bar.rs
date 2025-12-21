use crate::file_io::FileEncoding;
use crate::i18n::get_string;
use crate::line_column::calculate_line_column;
use std::sync::Mutex;
use std::sync::atomic::{AtomicBool, Ordering};
use windows::Win32::Foundation::COLORREF;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, CreateFontW, CreatePen, DRAW_TEXT_FORMAT, DeleteObject, DrawTextW, EndPaint,
    FONT_CHARSET, FONT_CLIP_PRECISION, FONT_OUTPUT_PRECISION, FONT_QUALITY, GetSysColor,
    GetSysColorBrush, HBRUSH, InvalidateRect, LineTo, MoveToEx, PAINTSTRUCT, PS_SOLID,
    SYS_COLOR_INDEX, SelectObject,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    DefWindowProcW, GetClientRect, GetWindowLongPtrW, IDC_ARROW, LoadCursorW, SendMessageW,
    SetCursor, SetWindowTextW, WINDOW_LONG_PTR_INDEX, WM_GETTEXT, WM_GETTEXTLENGTH, WM_PAINT,
    WM_SETCURSOR, WNDCLASS_STYLES, WNDCLASSW,
};
use windows::core::PCWSTR;

pub const EM_GETSEL: u32 = 0x00B0;
pub const EM_GETZOOM: u32 = 0x04E0;

// Helper function to convert raw SendMessageW result to i32
#[inline]
fn msg_as_i32(result: LRESULT) -> i32 {
    result.0 as i32
}

// Helper function to convert raw SendMessageW result to usize
#[inline]
fn msg_as_usize(result: LRESULT) -> usize {
    result.0 as usize
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
                if let Ok(cursor) = LoadCursorW(None, IDC_ARROW) {
                    SetCursor(Some(cursor));
                }
                LRESULT(1)
            }
            WM_PAINT => {
                let mut ps = PAINTSTRUCT::default();
                let hdc = BeginPaint(hwnd, &mut ps);

                let mut rect = RECT::default();
                let _ = GetClientRect(hwnd, &mut rect);

                // Gray color (RGB: 210, 209, 208)
                let separator_color = COLORREF(0x00D0D1D2u32);
                let pen = CreatePen(PS_SOLID, 1, separator_color);
                if !pen.is_invalid() {
                    let old_pen = SelectObject(hdc, pen.into());

                    let width = rect.right - rect.left;
                    let height = rect.bottom - rect.top;

                    if height > width {
                        // Vertical line - draw in the middle
                        let x = width / 2;
                        let _ = MoveToEx(hdc, x, rect.top, None);
                        let _ = LineTo(hdc, x, rect.bottom);
                    } else {
                        // Horizontal line - draw in the middle
                        let y = height / 2;
                        let _ = MoveToEx(hdc, rect.left, y, None);
                        let _ = LineTo(hdc, rect.right, y);
                    }

                    SelectObject(hdc, old_pen);
                    let _ = DeleteObject(pen.into());
                }

                let _ = EndPaint(hwnd, &ps);
                LRESULT(0)
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
                if let Ok(cursor) = LoadCursorW(None, IDC_ARROW) {
                    SetCursor(Some(cursor));
                }
                LRESULT(1)
            }
            WM_PAINT => {
                // Paint text using STATIC control behavior
                use windows::Win32::Graphics::Gdi::{
                    BACKGROUND_MODE, FillRect, SetBkMode, SetTextColor,
                };
                const GWL_STYLE: WINDOW_LONG_PTR_INDEX = WINDOW_LONG_PTR_INDEX(-16);

                let mut ps = PAINTSTRUCT::default();
                let hdc = BeginPaint(hwnd, &mut ps);

                let mut rect = RECT::default();
                let _ = GetClientRect(hwnd, &mut rect);

                // Fill background
                const COLOR_BTNFACE: SYS_COLOR_INDEX = SYS_COLOR_INDEX(15);
                let brush = GetSysColorBrush(COLOR_BTNFACE);
                FillRect(hdc, &rect, HBRUSH(brush.0));

                // Get window text
                let text_len =
                    SendMessageW(hwnd, WM_GETTEXTLENGTH, Some(WPARAM(0)), Some(LPARAM(0))).0
                        as usize;

                if text_len > 0 {
                    let mut buffer = vec![0u16; text_len + 1];
                    SendMessageW(
                        hwnd,
                        WM_GETTEXT,
                        Some(WPARAM(buffer.len())),
                        Some(LPARAM(buffer.as_mut_ptr() as isize)),
                    );

                    // Set text properties
                    let _ = SetBkMode(hdc, BACKGROUND_MODE(1)); // TRANSPARENT
                    const COLOR_BTNTEXT: SYS_COLOR_INDEX = SYS_COLOR_INDEX(18);
                    let _ = SetTextColor(hdc, COLORREF(GetSysColor(COLOR_BTNTEXT)));

                    // Create and set small font for status bar
                    let font_name = "Segoe UI";
                    let font = CreateFontW(
                        -12,                      // cHeight
                        0,                        // cWidth
                        0,                        // cEscapement
                        0,                        // cOrientation
                        400,                      // cWeight (FW_NORMAL)
                        0,                        // bItalic
                        0,                        // bUnderline
                        0,                        // bStrikeOut
                        FONT_CHARSET(1),          // iCharSet (DEFAULT_CHARSET)
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
                    if !font.is_invalid() {
                        let old_font = SelectObject(hdc, font.into());

                        // Check style for alignment
                        const SS_RIGHT: isize = 0x0002;
                        let style = GetWindowLongPtrW(hwnd, GWL_STYLE);
                        let format = if (style & SS_RIGHT) != 0 {
                            DRAW_TEXT_FORMAT(0x00000002 | 0x00000020 | 0x00000004) // DT_RIGHT | DT_SINGLELINE | DT_VCENTER
                        } else {
                            DRAW_TEXT_FORMAT(0x00000000 | 0x00000020 | 0x00000004) // DT_LEFT | DT_SINGLELINE | DT_VCENTER
                        };

                        let mut text_buffer = buffer[..text_len].to_vec();
                        let _ = DrawTextW(hdc, &mut text_buffer, &mut rect, format);

                        // Restore old font and delete created font
                        SelectObject(hdc, old_font);
                        let _ = DeleteObject(font.into());
                    }
                }

                let _ = EndPaint(hwnd, &ps);
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

// Register status bar window classes
pub unsafe fn register_status_bar_classes() {
    unsafe {
        let hinstance = GetModuleHandleW(None).unwrap_or_default();
        const COLOR_BTNFACE: SYS_COLOR_INDEX = SYS_COLOR_INDEX(15);

        // Register separator window class
        let separator_class_name = "SeparatorClass\0".encode_utf16().collect::<Vec<_>>();
        let separator_class = WNDCLASSW {
            style: WNDCLASS_STYLES(0x0001 | 0x0002), // CS_HREDRAW | CS_VREDRAW
            lpfnWndProc: Some(separator_proc),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: hinstance.into(),
            hIcon: Default::default(),
            hCursor: Default::default(),
            hbrBackground: HBRUSH(GetSysColorBrush(COLOR_BTNFACE).0),
            lpszMenuName: PCWSTR::null(),
            lpszClassName: PCWSTR(separator_class_name.as_ptr()),
        };
        let _ = windows::Win32::UI::WindowsAndMessaging::RegisterClassW(&separator_class);

        // Register status text window class
        let status_text_class_name = "StatusTextClass\0".encode_utf16().collect::<Vec<_>>();
        let status_text_class = WNDCLASSW {
            style: WNDCLASS_STYLES(0x0001 | 0x0002), // CS_HREDRAW | CS_VREDRAW
            lpfnWndProc: Some(status_text_proc),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: hinstance.into(),
            hIcon: Default::default(),
            hCursor: Default::default(),
            hbrBackground: HBRUSH(GetSysColorBrush(COLOR_BTNFACE).0),
            lpszMenuName: PCWSTR::null(),
            lpszClassName: PCWSTR(status_text_class_name.as_ptr()),
        };
        let _ = windows::Win32::UI::WindowsAndMessaging::RegisterClassW(&status_text_class);
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
            let mut start_pos: i32 = 0;
            let mut end_pos: i32 = 0;
            SendMessageW(
                edit_hwnd,
                EM_GETSEL,
                Some(WPARAM(&mut start_pos as *mut i32 as usize)),
                Some(LPARAM(&mut end_pos as *mut i32 as isize)),
            );

            // Get text length
            let text_length = msg_as_i32(SendMessageW(
                edit_hwnd,
                WM_GETTEXTLENGTH,
                Some(WPARAM(0)),
                Some(LPARAM(0)),
            ));

            let text_str = if text_length > 0 {
                // Allocate buffer and get text using WM_GETTEXT
                let mut buffer = vec![0u16; (text_length + 1) as usize];
                let actual_len = msg_as_usize(SendMessageW(
                    edit_hwnd,
                    WM_GETTEXT,
                    Some(WPARAM((text_length + 1) as usize)),
                    Some(LPARAM(buffer.as_mut_ptr() as isize)),
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

            // Convert char position to UTF-16 code unit position
            let mut utf16_pos = 0i32;
            let mut char_idx = 0i32;
            let mut chars_iter = text_str.chars().peekable();

            while let Some(ch) = chars_iter.next() {
                if char_idx >= start_pos {
                    break;
                }

                if ch == '\r' && chars_iter.peek() == Some(&'\n') {
                    chars_iter.next();
                    utf16_pos += 2;
                } else {
                    utf16_pos += if ch > '\u{FFFF}' { 2 } else { 1 };
                }

                char_idx += 1;
            }

            // Calculate line and column
            let (display_line, display_col) = calculate_line_column(&text_str, utf16_pos);

            // Get total character count
            let char_count = if COUNT_NEWLINE_AS_ONE.load(Ordering::SeqCst) {
                let mut count = 0i32;
                let mut chars = text_str.chars().peekable();
                while let Some(ch) = chars.next() {
                    if ch == '\r' && chars.peek() == Some(&'\n') {
                        chars.next();
                        count += 1;
                    } else if ch != '\0' {
                        count += 1;
                    }
                }
                count
            } else {
                text_length
            };

            // Get zoom level
            let mut numerator: i32 = 0;
            let mut denominator: i32 = 0;
            SendMessageW(
                edit_hwnd,
                EM_GETZOOM,
                Some(WPARAM(&mut numerator as *mut i32 as usize)),
                Some(LPARAM(&mut denominator as *mut i32 as isize)),
            );

            let zoom_percent = if denominator != 0 {
                (numerator * 100) / denominator
            } else {
                100
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
                let _ = SetWindowTextW(char_hwnd, PCWSTR(char_utf16.as_ptr()));
                let _ = InvalidateRect(Some(char_hwnd), None, true);

                // Update line and column
                let pos_format = get_string("STATUS_LINE_COL");
                let pos_text = format!(
                    "{}\0",
                    pos_format
                        .replace("{line}", &display_line.to_string())
                        .replace("{col}", &display_col.to_string())
                );
                let pos_utf16: Vec<u16> = pos_text.encode_utf16().collect();
                let _ = SetWindowTextW(pos_hwnd, PCWSTR(pos_utf16.as_ptr()));
                let _ = InvalidateRect(Some(pos_hwnd), None, true);

                // Update encoding display
                let encoding_text = match current_encoding {
                    FileEncoding::Utf8 => "UTF-8\0".to_string(),
                    FileEncoding::ShiftJis => "Shift-JIS\0".to_string(),
                    FileEncoding::Auto => format!("{}\0", get_string("ENCODING_AUTO")),
                };
                let encoding_utf16: Vec<u16> = encoding_text.encode_utf16().collect();
                let _ = SetWindowTextW(encoding_hwnd, PCWSTR(encoding_utf16.as_ptr()));
                let _ = InvalidateRect(Some(encoding_hwnd), None, true);

                // Update zoom level
                let zoom_text = format!("{}%\0", zoom_percent);
                let zoom_utf16: Vec<u16> = zoom_text.encode_utf16().collect();
                let _ = SetWindowTextW(zoom_hwnd, PCWSTR(zoom_utf16.as_ptr()));
                let _ = InvalidateRect(Some(zoom_hwnd), None, true);
            }
        }
    }
}
