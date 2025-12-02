use crate::i18n::get_string;
use std::fs;
use std::path::PathBuf;
use windows_sys::Win32::Foundation::HINSTANCE;
use windows_sys::Win32::UI::Controls::Dialogs::{
    GetOpenFileNameW, GetSaveFileNameW, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OPENFILENAMEW,
};
use windows_sys::Win32::UI::Input::KeyboardAndMouse::GetActiveWindow;

pub fn open_file_dialog() -> Option<PathBuf> {
    unsafe {
        let mut file_path: Vec<u16> = vec![0; 260];
        let all_files = get_string("FILE_FILTER_ALL");
        let text_files = get_string("FILE_FILTER_TEXT");
        let filter_text = format!("{}\0*.txt\0{}\0*.*\0\0", text_files, all_files);
        let filter = filter_text.encode_utf16().collect::<Vec<_>>();
        let title_text = format!("{}\0", get_string("FILE_DIALOG_OPEN"));
        let title = title_text.encode_utf16().collect::<Vec<_>>();
        let defext = "txt\0".encode_utf16().collect::<Vec<_>>();

        let mut ofn: OPENFILENAMEW = std::mem::zeroed();
        ofn.lStructSize = std::mem::size_of::<OPENFILENAMEW>() as u32;
        ofn.hwndOwner = GetActiveWindow();
        ofn.hInstance = HINSTANCE::default();
        ofn.lpstrFilter = filter.as_ptr();
        ofn.lpstrCustomFilter = std::ptr::null_mut();
        ofn.nMaxCustFilter = 0;
        ofn.nFilterIndex = 1; // Default to Text Files (*.txt)
        ofn.lpstrFile = file_path.as_mut_ptr();
        ofn.nMaxFile = file_path.len() as u32;
        ofn.lpstrFileTitle = std::ptr::null_mut();
        ofn.nMaxFileTitle = 0;
        ofn.lpstrInitialDir = std::ptr::null();
        ofn.lpstrTitle = title.as_ptr();
        ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
        ofn.nFileOffset = 0;
        ofn.nFileExtension = 0;
        ofn.lpstrDefExt = defext.as_ptr();
        ofn.lCustData = 0isize;
        ofn.lpfnHook = None;
        ofn.lpTemplateName = std::ptr::null();

        if GetOpenFileNameW(&mut ofn) != 0 {
            let path_len = file_path
                .iter()
                .position(|&c| c == 0)
                .unwrap_or(file_path.len());
            let path_str = String::from_utf16_lossy(&file_path[..path_len]);
            return Some(PathBuf::from(path_str.to_string()));
        }
        None
    }
}

pub fn save_file_dialog() -> Option<PathBuf> {
    unsafe {
        let mut file_path: Vec<u16> = vec![0; 260];
        let all_files = get_string("FILE_FILTER_ALL");
        let text_files = get_string("FILE_FILTER_TEXT");
        let filter_text = format!("{}\0*.txt\0{}\0*.*\0\0", text_files, all_files);
        let filter = filter_text.encode_utf16().collect::<Vec<_>>();
        let title_text = format!("{}\0", get_string("FILE_DIALOG_SAVE"));
        let title = title_text.encode_utf16().collect::<Vec<_>>();
        let defext = "txt\0".encode_utf16().collect::<Vec<_>>();

        let mut ofn: OPENFILENAMEW = std::mem::zeroed();
        ofn.lStructSize = std::mem::size_of::<OPENFILENAMEW>() as u32;
        ofn.hwndOwner = GetActiveWindow();
        ofn.hInstance = HINSTANCE::default();
        ofn.lpstrFilter = filter.as_ptr();
        ofn.lpstrCustomFilter = std::ptr::null_mut();
        ofn.nMaxCustFilter = 0;
        ofn.nFilterIndex = 1; // Default to Text Files (*.txt)
        ofn.lpstrFile = file_path.as_mut_ptr();
        ofn.nMaxFile = file_path.len() as u32;
        ofn.lpstrFileTitle = std::ptr::null_mut();
        ofn.nMaxFileTitle = 0;
        ofn.lpstrInitialDir = std::ptr::null();
        ofn.lpstrTitle = title.as_ptr();
        ofn.Flags = OFN_HIDEREADONLY;
        ofn.nFileOffset = 0;
        ofn.nFileExtension = 0;
        ofn.lpstrDefExt = defext.as_ptr();
        ofn.lCustData = 0isize;
        ofn.lpfnHook = None;
        ofn.lpTemplateName = std::ptr::null();

        if GetSaveFileNameW(&mut ofn) != 0 {
            let path_len = file_path
                .iter()
                .position(|&c| c == 0)
                .unwrap_or(file_path.len());
            let path_str = String::from_utf16_lossy(&file_path[..path_len]);
            return Some(PathBuf::from(path_str.to_string()));
        }
        None
    }
}

pub fn save_file(path: &PathBuf, content: &str) -> Result<(), Box<dyn std::error::Error>> {
    fs::write(path, content)?;
    Ok(())
}

pub fn load_file(path: &PathBuf) -> Result<String, Box<dyn std::error::Error>> {
    let content = fs::read_to_string(path)?;
    Ok(content)
}
