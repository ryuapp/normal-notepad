use crate::constants::{ID_EDIT_COPY, ID_EDIT_CUT, ID_EDIT_PASTE, ID_EDIT_SELECTALL};
use crate::i18n::get_string;
use windows_sys::Win32::Foundation::HWND;
use windows_sys::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CreatePopupMenu, DestroyMenu, MF_STRING, TrackPopupMenu,
};

/// Shows a context menu at the specified position
pub fn show_context_menu(hwnd: HWND, x: i32, y: i32) {
    unsafe {
        // Create popup menu
        let hmenu = CreatePopupMenu();
        if hmenu.is_null() {
            return;
        }

        // Get localized menu texts from i18n
        let cut_text = format!("{}\0", get_string("CONTEXT_CUT"));
        let copy_text = format!("{}\0", get_string("CONTEXT_COPY"));
        let paste_text = format!("{}\0", get_string("CONTEXT_PASTE"));
        let selectall_text = format!("{}\0", get_string("CONTEXT_SELECTALL"));

        // Convert to UTF-16
        let cut_utf16: Vec<u16> = cut_text.encode_utf16().collect();
        let copy_utf16: Vec<u16> = copy_text.encode_utf16().collect();
        let paste_utf16: Vec<u16> = paste_text.encode_utf16().collect();
        let selectall_utf16: Vec<u16> = selectall_text.encode_utf16().collect();

        // Add menu items
        AppendMenuW(hmenu, MF_STRING, ID_EDIT_CUT as usize, cut_utf16.as_ptr());
        AppendMenuW(hmenu, MF_STRING, ID_EDIT_COPY as usize, copy_utf16.as_ptr());
        AppendMenuW(
            hmenu,
            MF_STRING,
            ID_EDIT_PASTE as usize,
            paste_utf16.as_ptr(),
        );
        AppendMenuW(hmenu, 0x00000800, 0, std::ptr::null()); // MFT_SEPARATOR
        AppendMenuW(
            hmenu,
            MF_STRING,
            ID_EDIT_SELECTALL as usize,
            selectall_utf16.as_ptr(),
        );

        // Display popup menu
        TrackPopupMenu(hmenu, 0, x, y, 0, hwnd, std::ptr::null());

        // Clean up
        DestroyMenu(hmenu);
    }
}
