use serde::{Deserialize, Serialize};

/// Cell formatting style attributes (font color, background color, font styles, font family, font size).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
pub struct CellStyle {
    /// Font color as Hex (e.g. "#FF0000" or "FF0000") or standard color name ("red", "blue", etc.)
    pub font_color: Option<String>,
    /// Background fill color as Hex or color name
    pub bg_color: Option<String>,
    /// Bold text style flag
    pub bold: Option<bool>,
    /// Italic text style flag
    pub italic: Option<bool>,
    /// Underline text style flag
    pub underline: Option<bool>,
    /// Font family name (e.g. "Arial", "Calibri", "Courier New")
    pub font_family: Option<String>,
    /// Font size in points (e.g. 11, 12, 14)
    pub font_size: Option<u16>,
}

impl CellStyle {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn is_empty(&self) -> bool {
        self.font_color.is_none()
            && self.bg_color.is_none()
            && self.bold.is_none()
            && self.italic.is_none()
            && self.underline.is_none()
            && self.font_family.is_none()
            && self.font_size.is_none()
    }

    pub fn merge(&mut self, other: &CellStyle) {
        if other.font_color.is_some() {
            self.font_color = other.font_color.clone();
        }
        if other.bg_color.is_some() {
            self.bg_color = other.bg_color.clone();
        }
        if other.bold.is_some() {
            self.bold = other.bold;
        }
        if other.italic.is_some() {
            self.italic = other.italic;
        }
        if other.underline.is_some() {
            self.underline = other.underline;
        }
        if other.font_family.is_some() {
            self.font_family = other.font_family.clone();
        }
        if other.font_size.is_some() {
            self.font_size = other.font_size;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_style_is_empty() {
        let style = CellStyle::new();
        assert!(style.is_empty());

        let style_bold = CellStyle {
            bold: Some(true),
            ..Default::default()
        };
        assert!(!style_bold.is_empty());
    }

    #[test]
    fn test_cell_style_merge() {
        let mut style1 = CellStyle {
            font_color: Some("#FF0000".to_string()),
            bold: Some(true),
            ..Default::default()
        };
        let style2 = CellStyle {
            bg_color: Some("#00FF00".to_string()),
            font_color: Some("#0000FF".to_string()),
            font_size: Some(14),
            ..Default::default()
        };

        style1.merge(&style2);
        assert_eq!(style1.font_color, Some("#0000FF".to_string()));
        assert_eq!(style1.bg_color, Some("#00FF00".to_string()));
        assert_eq!(style1.bold, Some(true));
        assert_eq!(style1.font_size, Some(14));
    }
}
