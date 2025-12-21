use crate::constants::{
    ID_EDIT_COPY, ID_EDIT_CUT, ID_EDIT_DELETE, ID_EDIT_PASTE, ID_EDIT_REDO, ID_EDIT_SELECTALL,
    ID_EDIT_UNDO,
};
use crate::i18n::get_string;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::System::DataExchange::IsClipboardFormatAvailable;
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CreatePopupMenu, DestroyMenu, EnableMenuItem, GetWindowLongPtrW, MENU_ITEM_FLAGS,
    SendMessageW, TRACK_POPUP_MENU_FLAGS, TrackPopupMenu, WINDOW_LONG_PTR_INDEX,
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

        // Get edit control handle
        let edit_hwnd = HWND(GetWindowLongPtrW(hwnd, WINDOW_LONG_PTR_INDEX(0)) as _);

        // Check if there is a selection
        const EM_GETSEL: u32 = 0x00B0;
        let mut start_pos: i32 = 0;
        let mut end_pos: i32 = 0;
        SendMessageW(
            edit_hwnd,
            EM_GETSEL,
            Some(WPARAM(&mut start_pos as *mut i32 as usize)),
            Some(LPARAM(&mut end_pos as *mut i32 as isize)),
        );
        let has_selection = start_pos != end_pos;

        // Check if undo is available
        const EM_CANUNDO: u32 = 0x00C6;
        let can_undo = SendMessageW(edit_hwnd, EM_CANUNDO, Some(WPARAM(0)), Some(LPARAM(0))).0 != 0;

        // Check if redo is available
        const EM_CANREDO: u32 = 0x0455;
        let can_redo = SendMessageW(edit_hwnd, EM_CANREDO, Some(WPARAM(0)), Some(LPARAM(0))).0 != 0;

        // Check if paste is available
        const CF_UNICODETEXT: u32 = 13;
        let can_paste = IsClipboardFormatAvailable(CF_UNICODETEXT).is_ok();

        // Get localized menu texts from i18n
        let undo_text = format!("{}\0", get_string("CONTEXT_UNDO"));
        let redo_text = format!("{}\0", get_string("CONTEXT_REDO"));
        let cut_text = format!("{}\0", get_string("CONTEXT_CUT"));
        let copy_text = format!("{}\0", get_string("CONTEXT_COPY"));
        let paste_text = format!("{}\0", get_string("CONTEXT_PASTE"));
        let delete_text = format!("{}\0", get_string("CONTEXT_DELETE"));
        let selectall_text = format!("{}\0", get_string("CONTEXT_SELECTALL"));

        // Convert to UTF-16
        let undo_utf16: Vec<u16> = undo_text.encode_utf16().collect();
        let redo_utf16: Vec<u16> = redo_text.encode_utf16().collect();
        let cut_utf16: Vec<u16> = cut_text.encode_utf16().collect();
        let copy_utf16: Vec<u16> = copy_text.encode_utf16().collect();
        let paste_utf16: Vec<u16> = paste_text.encode_utf16().collect();
        let delete_utf16: Vec<u16> = delete_text.encode_utf16().collect();
        let selectall_utf16: Vec<u16> = selectall_text.encode_utf16().collect();

        // Add menu items
        let _ = AppendMenuW(
            hmenu,
            MENU_ITEM_FLAGS(0x00000000),
            ID_EDIT_UNDO as usize,
            PCWSTR(undo_utf16.as_ptr()),
        );
        let _ = AppendMenuW(
            hmenu,
            MENU_ITEM_FLAGS(0x00000000),
            ID_EDIT_REDO as usize,
            PCWSTR(redo_utf16.as_ptr()),
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
            MENU_ITEM_FLAGS(0x00000000),
            ID_EDIT_DELETE as usize,
            PCWSTR(delete_utf16.as_ptr()),
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

        // Enable/disable menu items based on state
        const MF_BYCOMMAND: u32 = 0x00000000;
        const MF_GRAYED: u32 = 0x00000001;

        // Disable undo if not available
        if !can_undo {
            let _ = EnableMenuItem(
                hmenu,
                ID_EDIT_UNDO as u32,
                MENU_ITEM_FLAGS(MF_BYCOMMAND | MF_GRAYED),
            );
        }

        // Disable redo if not available
        if !can_redo {
            let _ = EnableMenuItem(
                hmenu,
                ID_EDIT_REDO as u32,
                MENU_ITEM_FLAGS(MF_BYCOMMAND | MF_GRAYED),
            );
        }

        // Disable cut, copy and delete if no selection
        if !has_selection {
            let _ = EnableMenuItem(
                hmenu,
                ID_EDIT_CUT as u32,
                MENU_ITEM_FLAGS(MF_BYCOMMAND | MF_GRAYED),
            );
            let _ = EnableMenuItem(
                hmenu,
                ID_EDIT_COPY as u32,
                MENU_ITEM_FLAGS(MF_BYCOMMAND | MF_GRAYED),
            );
            let _ = EnableMenuItem(
                hmenu,
                ID_EDIT_DELETE as u32,
                MENU_ITEM_FLAGS(MF_BYCOMMAND | MF_GRAYED),
            );
        }

        // Disable paste if clipboard is empty
        if !can_paste {
            let _ = EnableMenuItem(
                hmenu,
                ID_EDIT_PASTE as u32,
                MENU_ITEM_FLAGS(MF_BYCOMMAND | MF_GRAYED),
            );
        }

        // Display popup menu
        let _ = TrackPopupMenu(hmenu, TRACK_POPUP_MENU_FLAGS(0), x, y, Some(0), hwnd, None);

        // Clean up
        let _ = DestroyMenu(hmenu);
    }
}
