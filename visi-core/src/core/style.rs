use serde::{Deserialize, Serialize};

#[doc = "OOXML cell border line style per ECMA-376 §18.18.3 ST_BorderStyle."]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default, Serialize, Deserialize)]
pub enum BorderStyle {
    #[doc = "No border line."]
    #[default]
    None,
    #[doc = "Thin solid line."]
    Thin,
    #[doc = "Medium solid line."]
    Medium,
    #[doc = "Thick solid line."]
    Thick,
    #[doc = "Dashed line."]
    Dashed,
    #[doc = "Dotted line."]
    Dotted,
    #[doc = "Double line."]
    Double,
    #[doc = "Hairline."]
    Hair,
    #[doc = "Medium dashed line."]
    MediumDashed,
    #[doc = "Dash-dot line."]
    DashDot,
    #[doc = "Medium dash-dot line."]
    MediumDashDot,
    #[doc = "Dash-dot-dot line."]
    DashDotDot,
    #[doc = "Medium dash-dot-dot line."]
    MediumDashDotDot,
    #[doc = "Slanted dash-dot line."]
    SlantDashDot,
}

impl BorderStyle {
    #[doc = "Returns the ECMA-376 ST_BorderStyle token."]
    pub fn as_ooxml(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Thin => "thin",
            Self::Medium => "medium",
            Self::Thick => "thick",
            Self::Dashed => "dashed",
            Self::Dotted => "dotted",
            Self::Double => "double",
            Self::Hair => "hair",
            Self::MediumDashed => "mediumDashed",
            Self::DashDot => "dashDot",
            Self::MediumDashDot => "mediumDashDot",
            Self::DashDotDot => "dashDotDot",
            Self::MediumDashDotDot => "mediumDashDotDot",
            Self::SlantDashDot => "slantDashDot",
        }
    }

    #[doc = "Parses an ECMA-376 ST_BorderStyle token."]
    pub fn from_ooxml(s: &str) -> Option<Self> {
        match s {
            "none" => Some(Self::None),
            "thin" => Some(Self::Thin),
            "medium" => Some(Self::Medium),
            "thick" => Some(Self::Thick),
            "dashed" => Some(Self::Dashed),
            "dotted" => Some(Self::Dotted),
            "double" => Some(Self::Double),
            "hair" => Some(Self::Hair),
            "mediumDashed" => Some(Self::MediumDashed),
            "dashDot" => Some(Self::DashDot),
            "mediumDashDot" => Some(Self::MediumDashDot),
            "dashDotDot" => Some(Self::DashDotDot),
            "mediumDashDotDot" => Some(Self::MediumDashDotDot),
            "slantDashDot" => Some(Self::SlantDashDot),
            _ => None,
        }
    }
}

#[doc = "A single edge of a cell border."]
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default, Serialize, Deserialize)]
pub struct BorderSide {
    #[doc = "Line style for this edge."]
    pub style: BorderStyle,
    #[doc = "Optional edge color as Hex (e.g. `#000000`) or standard color name."]
    #[serde(default)]
    pub color: Option<String>,
}

#[doc = "All border edges on a cell per ECMA-376 §18.8.4."]
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default, Serialize, Deserialize)]
pub struct CellBorders {
    #[doc = "Left edge."]
    #[serde(default)]
    pub left: Option<BorderSide>,
    #[doc = "Right edge."]
    #[serde(default)]
    pub right: Option<BorderSide>,
    #[doc = "Top edge."]
    #[serde(default)]
    pub top: Option<BorderSide>,
    #[doc = "Bottom edge."]
    #[serde(default)]
    pub bottom: Option<BorderSide>,
    #[doc = "Diagonal edge."]
    #[serde(default)]
    pub diagonal: Option<BorderSide>,
    #[doc = "Whether the diagonal runs bottom-left to top-right."]
    #[serde(default)]
    pub diagonal_up: bool,
    #[doc = "Whether the diagonal runs top-left to bottom-right."]
    #[serde(default)]
    pub diagonal_down: bool,
}

impl CellBorders {
    #[doc = "Returns true when no border edge is configured."]
    pub fn is_empty(&self) -> bool {
        self.left.is_none()
            && self.right.is_none()
            && self.top.is_none()
            && self.bottom.is_none()
            && self.diagonal.is_none()
    }
}

#[doc = "Horizontal cell alignment per ECMA-376 §18.18.40 ST_HorizontalAlignment."]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum HorizontalAlignment {
    #[doc = "General alignment."]
    General,
    #[doc = "Left aligned."]
    Left,
    #[doc = "Centered."]
    Center,
    #[doc = "Right aligned."]
    Right,
    #[doc = "Fill cell width."]
    Fill,
    #[doc = "Justified."]
    Justify,
    #[doc = "Center across selection."]
    CenterContinuous,
    #[doc = "Distributed."]
    Distributed,
}

impl HorizontalAlignment {
    #[doc = "Returns the ECMA-376 ST_HorizontalAlignment token."]
    pub fn as_ooxml(self) -> &'static str {
        match self {
            Self::General => "general",
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Fill => "fill",
            Self::Justify => "justify",
            Self::CenterContinuous => "centerContinuous",
            Self::Distributed => "distributed",
        }
    }

    #[doc = "Parses an ECMA-376 ST_HorizontalAlignment token."]
    pub fn from_ooxml(s: &str) -> Option<Self> {
        match s {
            "general" => Some(Self::General),
            "left" => Some(Self::Left),
            "center" => Some(Self::Center),
            "right" => Some(Self::Right),
            "fill" => Some(Self::Fill),
            "justify" => Some(Self::Justify),
            "centerContinuous" => Some(Self::CenterContinuous),
            "distributed" => Some(Self::Distributed),
            _ => None,
        }
    }
}

#[doc = "Vertical cell alignment per ECMA-376 §18.18.88 ST_VerticalAlignment."]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum VerticalAlignment {
    #[doc = "Top aligned."]
    Top,
    #[doc = "Vertically centered."]
    Center,
    #[doc = "Bottom aligned."]
    Bottom,
    #[doc = "Justified."]
    Justify,
    #[doc = "Distributed."]
    Distributed,
}

impl VerticalAlignment {
    #[doc = "Returns the ECMA-376 ST_VerticalAlignment token."]
    pub fn as_ooxml(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Bottom => "bottom",
            Self::Justify => "justify",
            Self::Distributed => "distributed",
        }
    }

    #[doc = "Parses an ECMA-376 ST_VerticalAlignment token."]
    pub fn from_ooxml(s: &str) -> Option<Self> {
        match s {
            "top" => Some(Self::Top),
            "center" => Some(Self::Center),
            "bottom" => Some(Self::Bottom),
            "justify" => Some(Self::Justify),
            "distributed" => Some(Self::Distributed),
            _ => None,
        }
    }
}

#[doc = "Cell text alignment and wrapping properties per ECMA-376 §18.8.1."]
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default, Serialize, Deserialize)]
pub struct CellAlignment {
    #[doc = "Horizontal alignment."]
    #[serde(default)]
    pub horizontal: Option<HorizontalAlignment>,
    #[doc = "Vertical alignment."]
    #[serde(default)]
    pub vertical: Option<VerticalAlignment>,
    #[doc = "Whether text wraps within the cell."]
    #[serde(default)]
    pub wrap_text: Option<bool>,
    #[doc = "Text rotation angle in degrees (0-180, or 255 for vertical)."]
    #[serde(default)]
    pub text_rotation: Option<u32>,
    #[doc = "Indent level."]
    #[serde(default)]
    pub indent: Option<u32>,
    #[doc = "Whether to shrink font size to fit cell width."]
    #[serde(default)]
    pub shrink_to_fit: Option<bool>,
}

impl CellAlignment {
    #[doc = "Returns true when no alignment attribute is set."]
    pub fn is_empty(&self) -> bool {
        self.horizontal.is_none()
            && self.vertical.is_none()
            && self.wrap_text.is_none()
            && self.text_rotation.is_none()
            && self.indent.is_none()
            && self.shrink_to_fit.is_none()
    }
}

#[doc = "Cell fill pattern per ECMA-376 §18.18.55 ST_PatternType."]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum FillPattern {
    #[doc = "No fill."]
    None,
    #[doc = "Solid fill."]
    Solid,
    #[doc = "Medium gray."]
    MediumGray,
    #[doc = "Dark gray."]
    DarkGray,
    #[doc = "Light gray."]
    LightGray,
    #[doc = "Dark horizontal stripes."]
    DarkHorizontal,
    #[doc = "Dark vertical stripes."]
    DarkVertical,
    #[doc = "Dark diagonal down stripes."]
    DarkDown,
    #[doc = "Dark diagonal up stripes."]
    DarkUp,
    #[doc = "Dark grid."]
    DarkGrid,
    #[doc = "Dark trellis."]
    DarkTrellis,
    #[doc = "Light horizontal stripes."]
    LightHorizontal,
    #[doc = "Light vertical stripes."]
    LightVertical,
    #[doc = "Light diagonal down stripes."]
    LightDown,
    #[doc = "Light diagonal up stripes."]
    LightUp,
    #[doc = "Light grid."]
    LightGrid,
    #[doc = "Light trellis."]
    LightTrellis,
    #[doc = "12.5% gray."]
    Gray125,
    #[doc = "6.25% gray."]
    Gray0625,
}

impl FillPattern {
    #[doc = "Returns the ECMA-376 ST_PatternType token."]
    pub fn as_ooxml(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Solid => "solid",
            Self::MediumGray => "mediumGray",
            Self::DarkGray => "darkGray",
            Self::LightGray => "lightGray",
            Self::DarkHorizontal => "darkHorizontal",
            Self::DarkVertical => "darkVertical",
            Self::DarkDown => "darkDown",
            Self::DarkUp => "darkUp",
            Self::DarkGrid => "darkGrid",
            Self::DarkTrellis => "darkTrellis",
            Self::LightHorizontal => "lightHorizontal",
            Self::LightVertical => "lightVertical",
            Self::LightDown => "lightDown",
            Self::LightUp => "lightUp",
            Self::LightGrid => "lightGrid",
            Self::LightTrellis => "lightTrellis",
            Self::Gray125 => "gray125",
            Self::Gray0625 => "gray0625",
        }
    }

    #[doc = "Parses an ECMA-376 ST_PatternType token."]
    pub fn from_ooxml(s: &str) -> Option<Self> {
        match s {
            "none" => Some(Self::None),
            "solid" => Some(Self::Solid),
            "mediumGray" => Some(Self::MediumGray),
            "darkGray" => Some(Self::DarkGray),
            "lightGray" => Some(Self::LightGray),
            "darkHorizontal" => Some(Self::DarkHorizontal),
            "darkVertical" => Some(Self::DarkVertical),
            "darkDown" => Some(Self::DarkDown),
            "darkUp" => Some(Self::DarkUp),
            "darkGrid" => Some(Self::DarkGrid),
            "darkTrellis" => Some(Self::DarkTrellis),
            "lightHorizontal" => Some(Self::LightHorizontal),
            "lightVertical" => Some(Self::LightVertical),
            "lightDown" => Some(Self::LightDown),
            "lightUp" => Some(Self::LightUp),
            "lightGrid" => Some(Self::LightGrid),
            "lightTrellis" => Some(Self::LightTrellis),
            "gray125" => Some(Self::Gray125),
            "gray0625" => Some(Self::Gray0625),
            _ => None,
        }
    }
}

#[doc = "Cell formatting style attributes."]
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct CellStyle {
    #[doc = "Font color as Hex (e.g. `#FF0000` or `FF0000`) or standard color name."]
    pub font_color: Option<String>,
    #[doc = "Background fill color as Hex or color name."]
    pub bg_color: Option<String>,
    #[doc = "Bold text style flag."]
    pub bold: Option<bool>,
    #[doc = "Italic text style flag."]
    pub italic: Option<bool>,
    #[doc = "Underline text style flag."]
    pub underline: Option<bool>,
    #[doc = "Strikethrough text style flag."]
    #[serde(default)]
    pub strikethrough: Option<bool>,
    #[doc = "Font family name (e.g. `Arial`, `Calibri`)."]
    pub font_family: Option<String>,
    #[doc = "Font size in points (e.g. 11, 12, 14.5)."]
    pub font_size: Option<f64>,
    #[doc = "Excel number-format code (e.g. `0.00`, `m/d/yy`)."]
    pub num_format: Option<String>,
    #[doc = "Fill pattern type; defaults to solid when `bg_color` is set."]
    #[serde(default)]
    pub fill_pattern: Option<FillPattern>,
    #[doc = "Cell borders."]
    #[serde(default)]
    pub borders: Option<CellBorders>,
    #[doc = "Cell alignment and wrapping."]
    #[serde(default)]
    pub alignment: Option<CellAlignment>,
}

impl CellStyle {
    #[doc = "A style with nothing set."]
    pub fn new() -> Self {
        Self::default()
    }

    #[doc = "Whether no attribute is set."]
    pub fn is_empty(&self) -> bool {
        self.font_color.is_none()
            && self.bg_color.is_none()
            && self.bold.is_none()
            && self.italic.is_none()
            && self.underline.is_none()
            && self.strikethrough.is_none()
            && self.font_family.is_none()
            && self.font_size.is_none()
            && self.num_format.is_none()
            && self.fill_pattern.is_none()
            && self.borders.as_ref().is_none_or(CellBorders::is_empty)
            && self.alignment.as_ref().is_none_or(CellAlignment::is_empty)
    }

    #[doc = "Overlays `other` onto this style, attribute by attribute."]
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
        if other.strikethrough.is_some() {
            self.strikethrough = other.strikethrough;
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
        if other.fill_pattern.is_some() {
            self.fill_pattern = other.fill_pattern;
        }
        if other.borders.as_ref().is_some_and(|b| !b.is_empty()) {
            self.borders = other.borders.clone();
        }
        if other.alignment.as_ref().is_some_and(|a| !a.is_empty()) {
            self.alignment = other.alignment.clone();
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
            strikethrough: Some(true),
            ..Default::default()
        };

        style1.merge(&style2);
        assert_eq!(style1.font_color, Some("#0000FF".to_string()));
        assert_eq!(style1.bg_color, Some("#00FF00".to_string()));
        assert_eq!(style1.bold, Some(true));
        assert_eq!(style1.font_size, Some(14.0));
        assert_eq!(style1.strikethrough, Some(true));
    }
}
