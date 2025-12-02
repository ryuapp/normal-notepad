use crate::i18n::get_string;
use crate::line_column::calculate_line_column;
use std::sync::Mutex;
use std::sync::atomic::{AtomicBool, Ordering};
use windows_sys::Win32::Foundation::HWND;
use windows_sys::Win32::UI::WindowsAndMessaging::SendMessageW;
use windows_sys::Win32::UI::WindowsAndMessaging::SetWindowTextW;

pub const EM_GETSEL: u32 = 0x00B0;
pub const EM_LINEFROMCHAR: u32 = 0x00C9;
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
static LAST_STATUS: Mutex<Option<(i32, i32, i32)>> = Mutex::new(None);

pub fn update_status_bar(edit_hwnd: HWND, char_hwnd: HWND, pos_hwnd: HWND) {
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

            // Check if values changed
            let current_status = (display_line, display_col, char_count);
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
            }
        }
    }
}
