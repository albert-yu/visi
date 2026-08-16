//! Per-cell formatting.

use serde::{Deserialize, Serialize};

/// Cell formatting style attributes (font color, background color, font styles, font family, font size).
// No `Eq`: `font_size` is an `f64`, matching Excel's `Double`-typed
// `Font.Size`. `PartialEq` is what the codebase actually uses.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
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
    /// Font size in points (e.g. 11, 12, 14.5).
    ///
    /// `f64` rather than an integer because Excel's is: `Font.Size` reports
    /// as a `Double` and a half-point size round-trips (`.Font.Size = 10.5`
    /// reads back as `10.5`), both measured with `fuzz/vba_style_probe.py`.
    pub font_size: Option<f64>,
    /// Excel number-format code (e.g. `m/d/yy`, `yyyy-mm-dd`).
    ///
    /// This is how a date cell remembers the notation it was written in: the
    /// value stays a plain numeric serial, exactly as in Excel, and the format
    /// governs only how it renders. See `core::date`.
    pub num_format: Option<String>,
}

impl CellStyle {
    /// A style with nothing set.
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether no attribute is set. An empty style is stored as no style at
    /// all rather than kept around.
    pub fn is_empty(&self) -> bool {
        self.font_color.is_none()
            && self.bg_color.is_none()
            && self.bold.is_none()
            && self.italic.is_none()
            && self.underline.is_none()
            && self.font_family.is_none()
            && self.font_size.is_none()
            && self.num_format.is_none()
    }

    /// Overlays `other` onto this style, attribute by attribute.
    ///
    /// Only the attributes `other` actually sets are copied, so merging a
    /// style that just sets `bold` leaves an existing font color alone.
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
        if other.num_format.is_some() {
            self.num_format = other.num_format.clone();
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
            font_size: Some(14.0),
            ..Default::default()
        };

        style1.merge(&style2);
        assert_eq!(style1.font_color, Some("#0000FF".to_string()));
        assert_eq!(style1.bg_color, Some("#00FF00".to_string()));
        assert_eq!(style1.bold, Some(true));
        assert_eq!(style1.font_size, Some(14.0));
    }
}
