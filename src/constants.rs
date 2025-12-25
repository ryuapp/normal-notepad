// Application-specific constants only
// Windows API constants should be imported directly from windows-sys in each file

// OLE constants
pub const OLE_PLACEHOLDER: u16 = 0x0001;

// Icon constants
pub const ICON_BIG: usize = 1;
pub const ICON_SMALL: usize = 0;

// Custom Windows API constants (not available in windows-sys)
pub const EM_EXLIMITTEXT: u32 = 0x0435;
pub const EM_SETTARGETDEVICE: u32 = 0x0448;
pub const EM_GETTEXT: u32 = 0x000D;
pub const EM_SETTEXT: u32 = 0x000C;
pub const EM_SETLANGOPTIONS: u32 = 0x0478;
pub const EM_GETLANGOPTIONS: u32 = 0x0479;
pub const EM_SETPARAFORMAT: u32 = 0x0447;
pub const ES_MULTILINE: u32 = 0x0004;
pub const EC_TOPMARGIN: u32 = 0x0002;
pub const IMF_AUTOFONT: u32 = 0x0002;
pub const IMF_DUALFONT: u32 = 0x0080;

// PARAFORMAT2 constants
pub const PFM_LINESPACING: u32 = 0x00000100;
pub const PFM_SPACEBEFORE: u32 = 0x00000040;
pub const PFM_SPACEAFTER: u32 = 0x00000080;

// Menu command IDs
pub const ID_FILE_NEW: i32 = 1;
pub const ID_FILE_OPEN: i32 = 2;
pub const ID_FILE_SAVE: i32 = 3;
pub const ID_FILE_SAVEAS: i32 = 4;
pub const ID_FILE_EXIT: i32 = 5;
pub const ID_EDIT_UNDO: i32 = 6;
pub const ID_EDIT_REDO: i32 = 7;
pub const ID_EDIT_COPY: i32 = 8;
pub const ID_EDIT_CUT: i32 = 9;
pub const ID_EDIT_PASTE: i32 = 10;
pub const ID_VIEW_WORDWRAP: i32 = 11;
pub const ID_VIEW_STATUSBAR: i32 = 12;
pub const ID_EDIT_SELECTALL: i32 = 13;
pub const ID_EDIT_DELETE: i32 = 14;
pub const ID_VIEW_DARKMODE: i32 = 15;
