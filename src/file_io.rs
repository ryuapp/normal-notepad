use crate::i18n::get_string;
use std::fs;
use std::path::PathBuf;
use windows::Win32::Globalization::{
    GetACP, MULTI_BYTE_TO_WIDE_CHAR_FLAGS, MultiByteToWideChar, WideCharToMultiByte,
};
use windows::Win32::System::Com::*;
use windows::Win32::UI::Shell::Common::*;
use windows::Win32::UI::Shell::*;
use windows::core::*;

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FileEncoding {
    Utf8,
    Utf8Bom,
    Utf16Le,
    Utf16Be,
    ShiftJis,
    Auto,
}

const ENCODING_CONTROL_ID: u32 = 2000;

pub fn open_file_dialog() -> Option<(PathBuf, FileEncoding)> {
    unsafe {
        // Initialize COM
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

        // Create FileOpenDialog
        let dialog: IFileOpenDialog = match CoCreateInstance(&FileOpenDialog, None, CLSCTX_ALL) {
            Ok(d) => d,
            Err(_) => {
                CoUninitialize();
                return None;
            }
        };

        // Set title
        let title = HSTRING::from(get_string("FILE_DIALOG_OPEN"));
        let _ = dialog.SetTitle(&title);

        // Set default extension
        let defext = w!("txt");
        let _ = dialog.SetDefaultExtension(defext);

        // Set file types
        let text_files = get_string("FILE_FILTER_TEXT");
        let all_files = get_string("FILE_FILTER_ALL");

        let text_files_hstring = HSTRING::from(&text_files);
        let all_files_hstring = HSTRING::from(&all_files);

        let file_types = [
            COMDLG_FILTERSPEC {
                pszName: PCWSTR(text_files_hstring.as_ptr()),
                pszSpec: w!("*.txt"),
            },
            COMDLG_FILTERSPEC {
                pszName: PCWSTR(all_files_hstring.as_ptr()),
                pszSpec: w!("*.*"),
            },
        ];

        let _ = dialog.SetFileTypes(&file_types);
        let _ = dialog.SetFileTypeIndex(1);

        // Get IFileDialogCustomize interface to add custom controls
        let customize: IFileDialogCustomize = match dialog.cast() {
            Ok(c) => c,
            Err(_) => {
                CoUninitialize();
                return None;
            }
        };

        // Add encoding controls with visual group title
        let group_id = ENCODING_CONTROL_ID + 100;
        let combo_id = ENCODING_CONTROL_ID + 101;

        // Get localized encoding label
        let encoding_label = get_string("FILE_ENCODING");
        let encoding_label_hstring = HSTRING::from(&encoding_label);
        let encoding_label_pcwstr = PCWSTR(encoding_label_hstring.as_ptr());

        // Start visual group with localized title
        if customize
            .StartVisualGroup(group_id, encoding_label_pcwstr)
            .is_ok()
        {
            // Add combo box
            if customize.AddComboBox(combo_id).is_ok() {
                let auto_text = get_string("ENCODING_AUTO");
                let auto_hstring = HSTRING::from(&auto_text);
                let auto_label = PCWSTR(auto_hstring.as_ptr());

                let ansi_text = get_string("ENCODING_ANSI");
                let ansi_hstring = HSTRING::from(&ansi_text);
                let ansi_label = PCWSTR(ansi_hstring.as_ptr());

                let utf16le_label = w!("UTF-16 LE");
                let utf16be_label = w!("UTF-16 BE");
                let utf8_label = w!("UTF-8");
                let utf8bom_label = w!("UTF-8 (BOM)");

                let _ = customize.AddControlItem(combo_id, 0, auto_label);
                let _ = customize.AddControlItem(combo_id, 1, ansi_label);
                let _ = customize.AddControlItem(combo_id, 2, utf16le_label);
                let _ = customize.AddControlItem(combo_id, 3, utf16be_label);
                let _ = customize.AddControlItem(combo_id, 4, utf8_label);
                let _ = customize.AddControlItem(combo_id, 5, utf8bom_label);
                let _ = customize.SetSelectedControlItem(combo_id, 0); // Default to Auto
            }

            let _ = customize.EndVisualGroup();

            // Make the group prominent
            let _ = customize.MakeProminent(group_id);
        }

        // Show dialog
        let mut encoding = FileEncoding::Auto;
        if dialog.Show(None).is_ok() {
            // Get selected encoding from combo box
            if let Ok(selected) = customize.GetSelectedControlItem(combo_id) {
                encoding = match selected {
                    1 => FileEncoding::ShiftJis,
                    2 => FileEncoding::Utf16Le,
                    3 => FileEncoding::Utf16Be,
                    4 => FileEncoding::Utf8,
                    5 => FileEncoding::Utf8Bom,
                    _ => FileEncoding::Auto,
                };
            }

            // Get file path
            if let Ok(result) = dialog.GetResult() {
                if let Ok(path) = result.GetDisplayName(SIGDN_FILESYSPATH) {
                    let path_str = path.to_string().ok()?;
                    CoTaskMemFree(Some(path.0 as _));
                    CoUninitialize();
                    return Some((PathBuf::from(path_str), encoding));
                }
            }
        }

        CoUninitialize();
        None
    }
}

pub fn save_file_dialog(
    default_encoding: FileEncoding,
    current_file: Option<&PathBuf>,
) -> Option<(PathBuf, FileEncoding)> {
    unsafe {
        // Initialize COM
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

        // Create FileSaveDialog
        let dialog: IFileSaveDialog = match CoCreateInstance(&FileSaveDialog, None, CLSCTX_ALL) {
            Ok(d) => d,
            Err(_) => {
                CoUninitialize();
                return None;
            }
        };

        // Set title
        let title = HSTRING::from(get_string("FILE_DIALOG_SAVE"));
        let _ = dialog.SetTitle(&title);

        // Set default extension
        let defext = w!("txt");
        let _ = dialog.SetDefaultExtension(defext);

        // Set file types
        let text_files = get_string("FILE_FILTER_TEXT");
        let all_files = get_string("FILE_FILTER_ALL");

        let text_files_hstring = HSTRING::from(&text_files);
        let all_files_hstring = HSTRING::from(&all_files);

        let file_types = [
            COMDLG_FILTERSPEC {
                pszName: PCWSTR(text_files_hstring.as_ptr()),
                pszSpec: w!("*.txt"),
            },
            COMDLG_FILTERSPEC {
                pszName: PCWSTR(all_files_hstring.as_ptr()),
                pszSpec: w!("*.*"),
            },
        ];

        let _ = dialog.SetFileTypes(&file_types);
        let _ = dialog.SetFileTypeIndex(1);

        // Set default filename if current_file is provided
        if let Some(path) = current_file {
            if let Some(filename) = path.file_name() {
                if let Some(filename_str) = filename.to_str() {
                    let filename_hstring = HSTRING::from(filename_str);
                    let _ = dialog.SetFileName(&filename_hstring);
                }
            }
        }

        // Get IFileDialogCustomize interface to add custom controls
        let customize: IFileDialogCustomize = match dialog.cast() {
            Ok(c) => c,
            Err(_) => {
                CoUninitialize();
                return None;
            }
        };

        // Add encoding controls with visual group title
        let group_id = ENCODING_CONTROL_ID;
        let combo_id = ENCODING_CONTROL_ID + 1;

        // Get localized encoding label
        let encoding_label = get_string("FILE_ENCODING");
        let encoding_label_hstring = HSTRING::from(&encoding_label);
        let encoding_label_pcwstr = PCWSTR(encoding_label_hstring.as_ptr());

        // Start visual group with localized title
        if customize
            .StartVisualGroup(group_id, encoding_label_pcwstr)
            .is_ok()
        {
            // Add combo box
            if customize.AddComboBox(combo_id).is_ok() {
                let ansi_text = get_string("ENCODING_ANSI");
                let ansi_hstring = HSTRING::from(&ansi_text);
                let ansi_label = PCWSTR(ansi_hstring.as_ptr());

                let utf16le_label = w!("UTF-16 LE");
                let utf16be_label = w!("UTF-16 BE");
                let utf8_label = w!("UTF-8");
                let utf8bom_label = w!("UTF-8 (BOM)");

                let _ = customize.AddControlItem(combo_id, 0, ansi_label);
                let _ = customize.AddControlItem(combo_id, 1, utf16le_label);
                let _ = customize.AddControlItem(combo_id, 2, utf16be_label);
                let _ = customize.AddControlItem(combo_id, 3, utf8_label);
                let _ = customize.AddControlItem(combo_id, 4, utf8bom_label);

                // Set default based on the provided encoding
                let default_index = match default_encoding {
                    FileEncoding::ShiftJis => 0,
                    FileEncoding::Utf16Le => 1,
                    FileEncoding::Utf16Be => 2,
                    FileEncoding::Utf8 | FileEncoding::Auto => 3,
                    FileEncoding::Utf8Bom => 4,
                };
                let _ = customize.SetSelectedControlItem(combo_id, default_index);
            }

            let _ = customize.EndVisualGroup();

            // Make the group prominent
            let _ = customize.MakeProminent(group_id);
        }

        // Show dialog
        let mut encoding = FileEncoding::Utf8;
        if dialog.Show(None).is_ok() {
            // Get selected encoding from combo box
            let combo_id = ENCODING_CONTROL_ID + 1;
            if let Ok(selected) = customize.GetSelectedControlItem(combo_id) {
                encoding = match selected {
                    0 => FileEncoding::ShiftJis,
                    1 => FileEncoding::Utf16Le,
                    2 => FileEncoding::Utf16Be,
                    4 => FileEncoding::Utf8Bom,
                    _ => FileEncoding::Utf8,
                };
            }

            // Get file path
            if let Ok(result) = dialog.GetResult() {
                if let Ok(path) = result.GetDisplayName(SIGDN_FILESYSPATH) {
                    let path_str = path.to_string().ok()?;
                    CoTaskMemFree(Some(path.0 as _));
                    CoUninitialize();
                    return Some((PathBuf::from(path_str), encoding));
                }
            }
        }

        CoUninitialize();
        None
    }
}

pub fn save_file(
    path: &PathBuf,
    content: &str,
    encoding: FileEncoding,
) -> std::result::Result<(), Box<dyn std::error::Error>> {
    match encoding {
        FileEncoding::Utf8 | FileEncoding::Auto => {
            fs::write(path, content)?;
        }
        FileEncoding::Utf8Bom => {
            let mut bytes = vec![0xEF, 0xBB, 0xBF];
            bytes.extend_from_slice(content.as_bytes());
            fs::write(path, &bytes)?;
        }
        FileEncoding::Utf16Le => {
            let mut bytes = vec![0xFF, 0xFE];
            let utf16: Vec<u16> = content.encode_utf16().collect();
            for &word in &utf16 {
                bytes.push((word & 0xFF) as u8);
                bytes.push((word >> 8) as u8);
            }
            fs::write(path, &bytes)?;
        }
        FileEncoding::Utf16Be => {
            let mut bytes = vec![0xFE, 0xFF];
            let utf16: Vec<u16> = content.encode_utf16().collect();
            for &word in &utf16 {
                bytes.push((word >> 8) as u8);
                bytes.push((word & 0xFF) as u8);
            }
            fs::write(path, &bytes)?;
        }
        FileEncoding::ShiftJis => unsafe {
            let utf16: Vec<u16> = content.encode_utf16().collect();
            let code_page = GetACP();

            let size = WideCharToMultiByte(code_page, 0, &utf16, None, None, None);

            if size > 0 {
                let mut buffer = vec![0u8; size as usize];
                WideCharToMultiByte(code_page, 0, &utf16, Some(&mut buffer), None, None);
                fs::write(path, &buffer)?;
            }
        },
    }
    Ok(())
}

pub fn load_file(
    path: &PathBuf,
    encoding: FileEncoding,
) -> std::result::Result<(String, FileEncoding), Box<dyn std::error::Error>> {
    match encoding {
        FileEncoding::Utf8 => {
            let content = fs::read_to_string(path)?;
            Ok((content, FileEncoding::Utf8))
        }
        FileEncoding::Utf8Bom => {
            let bytes = fs::read(path)?;
            let content = if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
                String::from_utf8(bytes[3..].to_vec())?
            } else {
                String::from_utf8(bytes)?
            };
            Ok((content, FileEncoding::Utf8Bom))
        }
        FileEncoding::Utf16Le => {
            let bytes = fs::read(path)?;
            let start = if bytes.starts_with(&[0xFF, 0xFE]) {
                2
            } else {
                0
            };
            let utf16_data: Vec<u16> = bytes[start..]
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect();
            let content = String::from_utf16_lossy(&utf16_data);
            Ok((content, FileEncoding::Utf16Le))
        }
        FileEncoding::Utf16Be => {
            let bytes = fs::read(path)?;
            let start = if bytes.starts_with(&[0xFE, 0xFF]) {
                2
            } else {
                0
            };
            let utf16_data: Vec<u16> = bytes[start..]
                .chunks_exact(2)
                .map(|chunk| u16::from_be_bytes([chunk[0], chunk[1]]))
                .collect();
            let content = String::from_utf16_lossy(&utf16_data);
            Ok((content, FileEncoding::Utf16Be))
        }
        FileEncoding::ShiftJis => {
            let bytes = fs::read(path)?;
            unsafe {
                let code_page = GetACP();

                let size =
                    MultiByteToWideChar(code_page, MULTI_BYTE_TO_WIDE_CHAR_FLAGS(0), &bytes, None);

                if size > 0 {
                    let mut buffer = vec![0u16; size as usize];
                    MultiByteToWideChar(
                        code_page,
                        MULTI_BYTE_TO_WIDE_CHAR_FLAGS(0),
                        &bytes,
                        Some(&mut buffer),
                    );
                    let content = String::from_utf16_lossy(&buffer);
                    Ok((content, FileEncoding::ShiftJis))
                } else {
                    Ok((String::new(), FileEncoding::ShiftJis))
                }
            }
        }
        FileEncoding::Auto => {
            let bytes = fs::read(path)?;

            // Check for UTF-16 LE BOM
            if bytes.starts_with(&[0xFF, 0xFE]) {
                let utf16_data: Vec<u16> = bytes[2..]
                    .chunks_exact(2)
                    .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                    .collect();
                let content = String::from_utf16_lossy(&utf16_data);
                return Ok((content, FileEncoding::Utf16Le));
            }

            // Check for UTF-16 BE BOM
            if bytes.starts_with(&[0xFE, 0xFF]) {
                let utf16_data: Vec<u16> = bytes[2..]
                    .chunks_exact(2)
                    .map(|chunk| u16::from_be_bytes([chunk[0], chunk[1]]))
                    .collect();
                let content = String::from_utf16_lossy(&utf16_data);
                return Ok((content, FileEncoding::Utf16Be));
            }

            // Check for UTF-8 BOM
            if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
                let content = String::from_utf8(bytes[3..].to_vec())?;
                return Ok((content, FileEncoding::Utf8Bom));
            }

            // Try UTF-8 first
            if let Ok(content) = String::from_utf8(bytes.clone()) {
                return Ok((content, FileEncoding::Utf8));
            }

            // Fall back to ANSI (system default code page)
            unsafe {
                let code_page = GetACP();

                let size =
                    MultiByteToWideChar(code_page, MULTI_BYTE_TO_WIDE_CHAR_FLAGS(0), &bytes, None);

                if size > 0 {
                    let mut buffer = vec![0u16; size as usize];
                    MultiByteToWideChar(
                        code_page,
                        MULTI_BYTE_TO_WIDE_CHAR_FLAGS(0),
                        &bytes,
                        Some(&mut buffer),
                    );
                    let content = String::from_utf16_lossy(&buffer);
                    Ok((content, FileEncoding::ShiftJis))
                } else {
                    Ok((String::new(), FileEncoding::ShiftJis))
                }
            }
        }
    }
}
