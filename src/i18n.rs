use std::sync::Mutex;
use windows::Win32::Globalization::GetUserDefaultUILanguage;

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Language {
    Japanese,
    English,
}

fn detect_system_language() -> Language {
    unsafe {
        let lang = GetUserDefaultUILanguage();
        // Japanese language ID is 0x11 (LANG_JAPANESE = 17)
        if (lang as u16) & 0xFF == 0x11 {
            Language::Japanese
        } else {
            Language::English
        }
    }
}

static CURRENT_LANGUAGE: Mutex<Language> = Mutex::new(Language::English);

pub fn init_language() {
    let detected_lang = detect_system_language();
    set_language(detected_lang);
}

pub fn get_language() -> Language {
    CURRENT_LANGUAGE
        .lock()
        .ok()
        .map(|l| *l)
        .unwrap_or(Language::English)
}

pub fn set_language(lang: Language) {
    if let Ok(mut current) = CURRENT_LANGUAGE.lock() {
        *current = lang;
    }
}

pub fn get_string(key: &str) -> String {
    match get_language() {
        Language::Japanese => get_japanese(key).to_string(),
        Language::English => get_english(key).to_string(),
    }
}

fn get_japanese(key: &str) -> &'static str {
    match key {
        // Menu
        "MENU_FILE" => "ファイル(&F)",
        "MENU_EDIT" => "編集(&E)",
        "MENU_VIEW" => "表示(&V)",
        "MENU_NEW" => "新規(&N)",
        "MENU_OPEN" => "開く(&O)",
        "MENU_SAVE" => "上書き保存(&S)",
        "MENU_SAVEAS" => "名前を付けて保存(&A)",
        "MENU_EXIT" => "終了(&X)",
        "MENU_UNDO" => "元に戻す(&U)",
        "MENU_REDO" => "やり直し(&R)",
        "MENU_COPY" => "コピー(&C)",
        "MENU_CUT" => "切り取り(&X)",
        "MENU_PASTE" => "貼り付け(&V)",
        "MENU_DELETE" => "削除(&D)",
        "MENU_SELECTALL" => "すべて選択(&A)",
        "MENU_WORDWRAP" => "右端で折り返す(&W)",
        "MENU_STATUSBAR" => "ステータスバー(&B)",
        "MENU_ZOOMIN" => "拡大(&I)",
        "MENU_ZOOMOUT" => "縮小(&O)",
        // Context menu
        "CONTEXT_UNDO" => "元に戻す (Ctrl+Z)",
        "CONTEXT_REDO" => "やり直し (Ctrl+Y)",
        "CONTEXT_CUT" => "切り取り (Ctrl+X)",
        "CONTEXT_COPY" => "コピー (Ctrl+C)",
        "CONTEXT_PASTE" => "貼り付け (Ctrl+V)",
        "CONTEXT_DELETE" => "削除 (Del)",
        "CONTEXT_SELECTALL" => "すべて選択 (Ctrl+A)",
        // Window title
        "WINDOW_TITLE" => "普通のメモ帳",
        // File
        "FILE_UNTITLED" => "無題",
        "FILE_DIALOG_OPEN" => "ファイルを開く",
        "FILE_DIALOG_SAVE" => "ファイルを保存",
        "FILE_FILTER_ALL" => "すべてのファイル (*.*)",
        "FILE_FILTER_TEXT" => "テキストファイル (*.txt)",
        "FILE_ENCODING" => "エンコード:",
        "ENCODING_AUTO" => "自動検出",
        "ENCODING_ANSI" => "Shift-JIS",
        // Status bar
        "STATUS_LINE_COL" => "行 {line}、列 {col}",
        "STATUS_CHAR_COUNT" => "{count} 文字",
        _ => "",
    }
}

fn get_english(key: &str) -> &'static str {
    match key {
        // Menu
        "MENU_FILE" => "File(&F)",
        "MENU_EDIT" => "Edit(&E)",
        "MENU_VIEW" => "View(&V)",
        "MENU_NEW" => "New(&N)",
        "MENU_OPEN" => "Open(&O)",
        "MENU_SAVE" => "Save(&S)",
        "MENU_SAVEAS" => "Save As(&A)",
        "MENU_EXIT" => "Exit(&X)",
        "MENU_UNDO" => "Undo(&U)",
        "MENU_REDO" => "Redo(&R)",
        "MENU_COPY" => "Copy(&C)",
        "MENU_CUT" => "Cut(&X)",
        "MENU_PASTE" => "Paste(&V)",
        "MENU_DELETE" => "Delete(&D)",
        "MENU_SELECTALL" => "Select All(&A)",
        "MENU_WORDWRAP" => "Word Wrap(&W)",
        "MENU_STATUSBAR" => "Status Bar(&B)",
        "MENU_ZOOMIN" => "Zoom In(&I)",
        "MENU_ZOOMOUT" => "Zoom Out(&O)",
        // Context menu
        "CONTEXT_UNDO" => "Undo (Ctrl+Z)",
        "CONTEXT_REDO" => "Redo (Ctrl+Y)",
        "CONTEXT_CUT" => "Cut (Ctrl+X)",
        "CONTEXT_COPY" => "Copy (Ctrl+C)",
        "CONTEXT_PASTE" => "Paste (Ctrl+V)",
        "CONTEXT_DELETE" => "Delete (Del)",
        "CONTEXT_SELECTALL" => "Select All (Ctrl+A)",
        // Window title
        "WINDOW_TITLE" => "Normal Notepad",
        // File
        "FILE_UNTITLED" => "Untitled",
        "FILE_DIALOG_OPEN" => "Open File",
        "FILE_DIALOG_SAVE" => "Save File",
        "FILE_FILTER_ALL" => "All Files (*.*)",
        "FILE_FILTER_TEXT" => "Text Files (*.txt)",
        "FILE_ENCODING" => "Encoding:",
        "ENCODING_AUTO" => "Auto",
        "ENCODING_ANSI" => "ANSI",
        // Status bar
        "STATUS_LINE_COL" => "Ln {line}, Col {col}",
        "STATUS_CHAR_COUNT" => "{count} characters",
        _ => "",
    }
}
