use crate::constants::{ID_EDIT_COPY, ID_EDIT_CUT, ID_EDIT_PASTE, ID_EDIT_SELECTALL};
use crate::i18n::get_string;
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CreatePopupMenu, DestroyMenu, MENU_ITEM_FLAGS, TRACK_POPUP_MENU_FLAGS,
    TrackPopupMenu,
};
use windows::core::PCWSTR;

/// Shows a context menu at the specified position
pub fn show_context_menu(hwnd: HWND, x: i32, y: i32) {
    unsafe {
        // Create popup menu
        let hmenu = CreatePopupMenu().ok();
        if hmenu.is_none() {
            return;
        }
        let hmenu = hmenu.unwrap();

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
        let _ = AppendMenuW(
            hmenu,
            MENU_ITEM_FLAGS(0x00000000),
            ID_EDIT_CUT as usize,
            PCWSTR(cut_utf16.as_ptr()),
        );
        let _ = AppendMenuW(
            hmenu,
            MENU_ITEM_FLAGS(0x00000000),
            ID_EDIT_COPY as usize,
            PCWSTR(copy_utf16.as_ptr()),
        );
        let _ = AppendMenuW(
            hmenu,
            MENU_ITEM_FLAGS(0x00000000),
            ID_EDIT_PASTE as usize,
            PCWSTR(paste_utf16.as_ptr()),
        );
        let _ = AppendMenuW(
            hmenu,
            MENU_ITEM_FLAGS(0x00000800), // MF_SEPARATOR
            0,
            PCWSTR::null(),
        );
        let _ = AppendMenuW(
            hmenu,
            MENU_ITEM_FLAGS(0x00000000),
            ID_EDIT_SELECTALL as usize,
            PCWSTR(selectall_utf16.as_ptr()),
        );

        // Display popup menu
        let _ = TrackPopupMenu(hmenu, TRACK_POPUP_MENU_FLAGS(0), x, y, Some(0), hwnd, None);

        // Clean up
        let _ = DestroyMenu(hmenu);
    }
}
