use crate::i18n::get_string;
use std::fs;
use std::path::PathBuf;
use windows::Win32::System::Com::*;
use windows::Win32::UI::Shell::Common::*;
use windows::Win32::UI::Shell::*;
use windows::core::*;

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FileEncoding {
    Utf8,
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
                // Get localized auto label
                let auto_text = get_string("ENCODING_AUTO");
                let auto_hstring = HSTRING::from(&auto_text);
                let auto_label = PCWSTR(auto_hstring.as_ptr());

                let utf8_label = w!("UTF-8");
                let sjis_label = w!("Shift-JIS");

                let _ = customize.AddControlItem(combo_id, 0, auto_label);
                let _ = customize.AddControlItem(combo_id, 1, utf8_label);
                let _ = customize.AddControlItem(combo_id, 2, sjis_label);
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
                    1 => FileEncoding::Utf8,
                    2 => FileEncoding::ShiftJis,
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
                let utf8_label = w!("UTF-8");
                let sjis_label = w!("Shift-JIS");

                let _ = customize.AddControlItem(combo_id, 0, utf8_label);
                let _ = customize.AddControlItem(combo_id, 1, sjis_label);

                // Set default based on the provided encoding
                let default_index = match default_encoding {
                    FileEncoding::ShiftJis => 1,
                    _ => 0, // UTF-8 for both Utf8 and Auto
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
                encoding = if selected == 1 {
                    FileEncoding::ShiftJis
                } else {
                    FileEncoding::Utf8
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
            // Auto defaults to UTF-8 for saving
            fs::write(path, content)?;
        }
        FileEncoding::ShiftJis => {
            // For Shift-JIS, we need to convert the string
            // Using Windows-31J (CP932) which is the Windows variant of Shift-JIS
            use encoding_rs::SHIFT_JIS;
            let (encoded, _, _) = SHIFT_JIS.encode(content);
            fs::write(path, encoded.as_ref())?;
        }
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
        FileEncoding::ShiftJis => {
            use encoding_rs::SHIFT_JIS;
            let bytes = fs::read(path)?;
            let (decoded, _, _) = SHIFT_JIS.decode(&bytes);
            Ok((decoded.into_owned(), FileEncoding::ShiftJis))
        }
        FileEncoding::Auto => {
            let bytes = fs::read(path)?;

            // Check for BOM
            if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
                // UTF-8 with BOM
                let content = String::from_utf8(bytes[3..].to_vec())?;
                return Ok((content, FileEncoding::Utf8));
            }

            // Try UTF-8 first
            if let Ok(content) = String::from_utf8(bytes.clone()) {
                return Ok((content, FileEncoding::Utf8));
            }

            // Fall back to Shift-JIS
            use encoding_rs::SHIFT_JIS;
            let (decoded, _, _) = SHIFT_JIS.decode(&bytes);
            Ok((decoded.into_owned(), FileEncoding::ShiftJis))
        }
    }
}
