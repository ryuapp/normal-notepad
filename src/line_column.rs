/// Calculate line and column number from cursor position in text
///
/// # Arguments
/// * `text` - The full text content (as UTF-16 would be)
/// * `cursor_pos` - Cursor position in UTF-16 code units
///
/// # Returns
/// Tuple of (line_number, column_number) where both are 1-indexed
pub fn calculate_line_column(text: &str, cursor_pos: i32) -> (i32, i32) {
    let mut line = 1;
    let mut col = 1;
    let mut utf16_pos = 0;
    let mut chars_iter = text.chars().peekable();

    while let Some(ch) = chars_iter.next() {
        // Check if we've reached the cursor position
        if utf16_pos >= cursor_pos {
            break;
        }

        let ch_utf16_len = if ch > '\u{FFFF}' { 2 } else { 1 };

        if ch == '\r' {
            // CR - line break
            utf16_pos += ch_utf16_len;

            // Skip LF if it follows CR
            if chars_iter.peek() == Some(&'\n') {
                chars_iter.next();
                utf16_pos += 1;
            }

            // Advance to the next line
            line += 1;
            col = 1;

            // Check if cursor is on the CR or LF
            if utf16_pos >= cursor_pos {
                break;
            }
        } else if ch == '\n' {
            // Standalone LF - also a line break
            utf16_pos += ch_utf16_len;

            // Advance to the next line
            line += 1;
            col = 1;

            // Check if cursor is on the LF
            if utf16_pos >= cursor_pos {
                break;
            }
        } else if ch != '\0' {
            // Regular character
            utf16_pos += ch_utf16_len;

            // We count this character because the cursor is AT or AFTER it
            col += 1;

            // Now check if we've reached the cursor position
            if utf16_pos >= cursor_pos {
                break;
            }
        } else {
            // Null terminator
            utf16_pos += ch_utf16_len;
        }
    }

    // Return values (already 1-indexed)
    (line, col)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_first_line_first_column() {
        let text = "hello world";
        let (line, col) = calculate_line_column(text, 0);
        assert_eq!((line, col), (1, 1));
    }

    #[test]
    fn test_first_line_middle() {
        let text = "hello world";
        let (line, col) = calculate_line_column(text, 5);
        assert_eq!((line, col), (1, 6));
    }

    #[test]
    fn test_second_line_first_column() {
        let text = "hello\r\nworld";
        let (line, col) = calculate_line_column(text, 7);
        assert_eq!((line, col), (2, 1));
    }

    #[test]
    fn test_second_line_middle() {
        let text = "hello\r\nworld";
        let (line, col) = calculate_line_column(text, 9);
        assert_eq!((line, col), (2, 3));
    }

    #[test]
    fn test_multiple_lines() {
        let text = "line1\r\nline2\r\nline3";
        let (line, col) = calculate_line_column(text, 13);
        assert_eq!((line, col), (3, 1));
    }

    #[test]
    fn test_multiple_lines_middle() {
        let text = "line1\r\nline2\r\nline3";
        let (line, col) = calculate_line_column(text, 16);
        assert_eq!((line, col), (3, 3));
    }

    #[test]
    fn test_empty_line() {
        let text = "line1\r\n\r\nline3";
        let (line, col) = calculate_line_column(text, 7);
        assert_eq!((line, col), (2, 1));
    }

    #[test]
    fn test_japanese_text() {
        let text = "あいう\r\nえお";
        let (line, col) = calculate_line_column(text, 4);
        assert_eq!((line, col), (2, 1));
    }

    #[test]
    fn test_japanese_text_middle() {
        let text = "あいう\r\nえお";
        let (line, col) = calculate_line_column(text, 6);
        assert_eq!((line, col), (2, 2));
    }

    #[test]
    fn test_end_of_text() {
        let text = "hello\r\nworld";
        let (line, col) = calculate_line_column(text, 12);
        assert_eq!((line, col), (2, 6));
    }
}
