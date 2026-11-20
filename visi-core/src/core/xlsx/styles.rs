use std::collections::HashMap;

use quick_xml::events::Event;

use crate::core::CellStyle;

pub(crate) const BUILTIN_NUM_FMTS: &[(u32, &str)] = &[
    (0, "General"),
    (1, "0"),
    (2, "0.00"),
    (3, "#,##0"),
    (4, "#,##0.00"),
    (5, "\"$\"#,##0_);(\"$\"#,##0)"),
    (6, "\"$\"#,##0_);[Red](\"$\"#,##0)"),
    (7, "\"$\"#,##0.00_);(\"$\"#,##0.00)"),
    (8, "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)"),
    (9, "0%"),
    (10, "0.00%"),
    (11, "0.00E+00"),
    (12, "# ?/?"),
    (13, "# ??/??"),
    (14, "m/d/yy"),
    (15, "d-mmm-yy"),
    (16, "d-mmm"),
    (17, "mmm-yy"),
    (18, "h:mm AM/PM"),
    (19, "h:mm:ss AM/PM"),
    (20, "h:mm"),
    (21, "h:mm:ss"),
    (22, "m/d/yy h:mm"),
    (37, "#,##0;(#,##0)"),
    (38, "#,##0;[Red](#,##0)"),
    (39, "#,##0.00;(#,##0.00)"),
    (40, "#,##0.00;[Red](#,##0.00)"),
    (41, r#"_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)"#),
    (42, r#"_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)"#),
    (43, r#"_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)"#),
    (
        44,
        r#"_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)"#,
    ),
    (45, "mm:ss"),
    (46, "[h]:mm:ss"),
    (47, "mm:ss.0"),
    (48, "##0.0E+0"),
    (49, "@"),
];

pub(crate) fn builtin_num_fmt_code(id: u32) -> Option<&'static str> {
    BUILTIN_NUM_FMTS
        .iter()
        .find(|(builtin_id, _)| *builtin_id == id)
        .map(|(_, code)| *code)
}

#[derive(Debug, Default)]
pub(crate) struct ParsedStyles {
    pub(crate) xf_styles: HashMap<u32, CellStyle>,
    pub(crate) xf_is_date: Vec<bool>,
}

impl ParsedStyles {
    pub(crate) fn is_date_xf(&self, xf_idx: u32) -> bool {
        self.xf_is_date
            .get(xf_idx as usize)
            .copied()
            .unwrap_or(false)
    }
}

pub(crate) fn parse_styles(xml: &str) -> ParsedStyles {
    let mut custom_num_formats: HashMap<u32, String> = HashMap::new();
    let mut fonts: Vec<CellStyle> = Vec::new();
    let mut fills: Vec<Option<String>> = Vec::new();
    let mut xf_styles: HashMap<u32, CellStyle> = HashMap::new();
    let mut xf_is_date: Vec<bool> = Vec::new();

    let mut in_fonts = false;
    let mut in_font = false;
    let mut current_font = CellStyle::default();
    let mut in_fills = false;
    let mut in_fill = false;
    let mut current_fill: Option<String> = None;
    let mut in_cell_xfs = false;
    let mut xf_idx = 0u32;

    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"numFmt" => {
                    if let Some(id) =
                        super::get_attr(e, b"numFmtId").and_then(|s| s.parse::<u32>().ok())
                        && let Some(code) = super::get_attr(e, b"formatCode")
                    {
                        custom_num_formats.insert(id, code);
                    }
                }
                b"fonts" => in_fonts = true,
                b"font" if in_fonts => {
                    in_font = true;
                    current_font = CellStyle::default();
                }
                b"b" if in_font => {
                    if style_bool_enabled(e) {
                        current_font.bold = Some(true);
                    }
                }
                b"i" if in_font => {
                    if style_bool_enabled(e) {
                        current_font.italic = Some(true);
                    }
                }
                b"u" if in_font => {
                    if style_bool_enabled(e) {
                        current_font.underline = Some(true);
                    }
                }
                b"color" if in_font => current_font.font_color = style_color_hex(e),
                b"name" if in_font => current_font.font_family = super::get_attr(e, b"val"),
                b"sz" if in_font => {
                    current_font.font_size =
                        super::get_attr(e, b"val").and_then(|s| s.parse().ok());
                }
                b"fills" => in_fills = true,
                b"fill" if in_fills => {
                    in_fill = true;
                    current_fill = None;
                }
                b"fgColor" if in_fill => {
                    if let Some(color) = style_color_hex(e) {
                        current_fill = Some(color);
                    }
                }
                b"bgColor" if in_fill => {
                    if current_fill.is_none() {
                        current_fill = style_color_hex(e);
                    }
                }
                b"cellXfs" => in_cell_xfs = true,
                b"xf" if in_cell_xfs => {
                    let mut style = CellStyle::default();
                    if let Some(font_id) =
                        super::get_attr(e, b"fontId").and_then(|s| s.parse::<usize>().ok())
                        && font_id != 0
                        && let Some(font) = fonts.get(font_id)
                    {
                        style.merge(font);
                    }
                    if let Some(fill_id) =
                        super::get_attr(e, b"fillId").and_then(|s| s.parse::<usize>().ok())
                        && fill_id > 1
                        && let Some(Some(bg_color)) = fills.get(fill_id)
                    {
                        style.bg_color = Some(bg_color.clone());
                    }
                    let mut is_date = false;
                    if let Some(num_fmt_id) =
                        super::get_attr(e, b"numFmtId").and_then(|s| s.parse::<u32>().ok())
                    {
                        let code = custom_num_formats
                            .get(&num_fmt_id)
                            .cloned()
                            .or_else(|| builtin_num_fmt_code(num_fmt_id).map(str::to_string));
                        if let Some(ref c) = code {
                            is_date = matches!(num_fmt_id, 14..=22 | 45..=47)
                                || crate::core::date::is_date_code(c);
                        }
                        style.num_format = code.filter(|c| !c.eq_ignore_ascii_case("General"));
                    }
                    if !style.is_empty() {
                        xf_styles.insert(xf_idx, style);
                    }
                    xf_is_date.push(is_date);
                    xf_idx += 1;
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"fonts" => in_fonts = false,
                b"font" if in_font => {
                    fonts.push(current_font.clone());
                    current_font = CellStyle::default();
                    in_font = false;
                }
                b"fills" => in_fills = false,
                b"fill" if in_fill => {
                    fills.push(current_fill.clone());
                    current_fill = None;
                    in_fill = false;
                }
                b"cellXfs" => in_cell_xfs = false,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    ParsedStyles {
        xf_styles,
        xf_is_date,
    }
}

fn style_bool_enabled(e: &quick_xml::events::BytesStart) -> bool {
    super::get_attr(e, b"val").is_none_or(|value| {
        value != "0" && !value.eq_ignore_ascii_case("false") && !value.eq_ignore_ascii_case("none")
    })
}

fn style_color_hex(e: &quick_xml::events::BytesStart) -> Option<String> {
    if let Some(rgb) = super::get_attr(e, b"rgb") {
        let trimmed = rgb.trim().trim_start_matches('#');
        let hex = match trimmed.len() {
            8 => &trimmed[2..],
            6 => trimmed,
            _ => return None,
        };
        return Some(format!("#{}", hex.to_ascii_uppercase()));
    }
    super::get_attr(e, b"indexed")
        .and_then(|idx| idx.parse::<u32>().ok())
        .and_then(indexed_color_hex)
}

fn indexed_color_hex(index: u32) -> Option<String> {
    let hex = match index {
        0 | 8 => "#000000",
        1 | 9 => "#FFFFFF",
        2 | 10 => "#FF0000",
        3 | 11 => "#00FF00",
        4 | 12 => "#0000FF",
        5 | 13 => "#FFFF00",
        6 | 14 => "#FF00FF",
        7 | 15 => "#00FFFF",
        _ => return None,
    };
    Some(hex.to_string())
}

pub(crate) fn parse_color_argb(color_str: &str) -> Option<String> {
    let s = color_str.trim().to_ascii_lowercase();
    let hex = match s.as_str() {
        "black" => "000000",
        "white" => "FFFFFF",
        "red" => "FF0000",
        "blue" => "0000FF",
        "green" => "008000",
        "yellow" => "FFFF00",
        "gray" | "grey" => "808080",
        _ => {
            let trimmed = s.trim_start_matches('#');
            if trimmed.len() == 6 && trimmed.chars().all(|c| c.is_ascii_hexdigit()) {
                return Some(format!("FF{}", trimmed.to_ascii_uppercase()));
            }
            if trimmed.len() == 8 && trimmed.chars().all(|c| c.is_ascii_hexdigit()) {
                return Some(trimmed.to_ascii_uppercase());
            }
            return None;
        }
    };
    Some(format!("FF{hex}"))
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct FontKey {
    bold: bool,
    italic: bool,
    underline: bool,
    size_tenths: Option<u32>,
    color_argb: Option<String>,
    family: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct XfKey {
    num_fmt_id: u32,
    font_id: usize,
    fill_id: usize,
}

#[derive(Debug)]
pub(crate) struct StyleRegistry {
    custom_num_fmts: Vec<(u32, String)>,
    num_fmt_map: HashMap<String, u32>,
    next_num_fmt_id: u32,
    fonts: Vec<FontKey>,
    font_map: HashMap<FontKey, usize>,
    fills: Vec<Option<String>>,
    fill_map: HashMap<String, usize>,
    xfs: Vec<XfKey>,
    xf_map: HashMap<XfKey, u32>,
}

impl Default for StyleRegistry {
    fn default() -> Self {
        let mut num_fmt_map = HashMap::new();
        for &(id, code) in BUILTIN_NUM_FMTS {
            num_fmt_map.insert(code.to_string(), id);
        }
        let default_font = FontKey {
            bold: false,
            italic: false,
            underline: false,
            size_tenths: Some(110),
            color_argb: None,
            family: Some("Calibri".to_string()),
        };
        let mut font_map = HashMap::new();
        font_map.insert(default_font.clone(), 0);

        let default_xf = XfKey {
            num_fmt_id: 0,
            font_id: 0,
            fill_id: 0,
        };
        let mut xf_map = HashMap::new();
        xf_map.insert(default_xf.clone(), 0);

        Self {
            custom_num_fmts: Vec::new(),
            num_fmt_map,
            next_num_fmt_id: 164,
            fonts: vec![default_font],
            font_map,
            fills: vec![None, None],
            fill_map: HashMap::new(),
            xfs: vec![default_xf],
            xf_map,
        }
    }
}

impl StyleRegistry {
    pub(crate) fn register_style(&mut self, style: &CellStyle) -> u32 {
        if style.is_empty() {
            return 0;
        }
        let num_fmt_id = match style.num_format.as_deref() {
            Some(code) if !code.is_empty() && !code.eq_ignore_ascii_case("General") => {
                if let Some(&id) = self.num_fmt_map.get(code) {
                    id
                } else {
                    let id = self.next_num_fmt_id;
                    self.next_num_fmt_id += 1;
                    self.custom_num_fmts.push((id, code.to_string()));
                    self.num_fmt_map.insert(code.to_string(), id);
                    id
                }
            }
            _ => 0,
        };

        let has_font_override = style.bold.unwrap_or(false)
            || style.italic.unwrap_or(false)
            || style.underline.unwrap_or(false)
            || style.font_color.is_some()
            || style.font_family.is_some()
            || style.font_size.is_some();

        let font_id = if has_font_override {
            let key = FontKey {
                bold: style.bold.unwrap_or(false),
                italic: style.italic.unwrap_or(false),
                underline: style.underline.unwrap_or(false),
                size_tenths: style
                    .font_size
                    .filter(|s| s.is_finite() && *s > 0.0)
                    .map(|s| (s * 10.0).round() as u32),
                color_argb: style.font_color.as_deref().and_then(parse_color_argb),
                family: style.font_family.clone(),
            };
            if let Some(&id) = self.font_map.get(&key) {
                id
            } else {
                let id = self.fonts.len();
                self.fonts.push(key.clone());
                self.font_map.insert(key, id);
                id
            }
        } else {
            0
        };

        let fill_id = match style.bg_color.as_deref().and_then(parse_color_argb) {
            Some(argb) => {
                if let Some(&id) = self.fill_map.get(&argb) {
                    id
                } else {
                    let id = self.fills.len();
                    self.fills.push(Some(argb.clone()));
                    self.fill_map.insert(argb, id);
                    id
                }
            }
            None => 0,
        };

        let xf_key = XfKey {
            num_fmt_id,
            font_id,
            fill_id,
        };
        if let Some(&idx) = self.xf_map.get(&xf_key) {
            idx
        } else {
            let idx = self.xfs.len() as u32;
            self.xfs.push(xf_key.clone());
            self.xf_map.insert(xf_key, idx);
            idx
        }
    }

    pub(crate) fn to_xml(&self, flavor: crate::core::xlsx::xml_util::NamespaceFlavor) -> String {
        let ns = flavor.spreadsheetml();
        let mut xml = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"{ns}\">"
        );

        if !self.custom_num_fmts.is_empty() {
            xml.push_str(&format!(
                "<numFmts count=\"{}\">",
                self.custom_num_fmts.len()
            ));
            for (id, code) in &self.custom_num_fmts {
                xml.push_str(&format!(
                    "<numFmt numFmtId=\"{id}\" formatCode=\"{}\"/>",
                    super::escape_xml(code)
                ));
            }
            xml.push_str("</numFmts>");
        }

        xml.push_str(&format!("<fonts count=\"{}\">", self.fonts.len()));
        for font in &self.fonts {
            xml.push_str("<font>");
            if font.bold {
                xml.push_str("<b/>");
            }
            if font.italic {
                xml.push_str("<i/>");
            }
            if font.underline {
                xml.push_str("<u/>");
            }
            let sz = font.size_tenths.map(|t| t as f64 / 10.0).unwrap_or(11.0);
            xml.push_str(&format!("<sz val=\"{sz}\"/>"));
            if let Some(ref argb) = font.color_argb {
                xml.push_str(&format!("<color rgb=\"{argb}\"/>"));
            }
            let family = font.family.as_deref().unwrap_or("Calibri");
            xml.push_str(&format!("<name val=\"{}\"/>", super::escape_xml(family)));
            xml.push_str("</font>");
        }
        xml.push_str("</fonts>");

        xml.push_str(&format!("<fills count=\"{}\">", self.fills.len()));
        xml.push_str("<fill><patternFill patternType=\"none\"/></fill>");
        xml.push_str("<fill><patternFill patternType=\"gray125\"/></fill>");
        for fill in self.fills.iter().skip(2) {
            if let Some(argb) = fill {
                xml.push_str(&format!(
                    "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{argb}\"/><bgColor indexed=\"64\"/></patternFill></fill>"
                ));
            }
        }
        xml.push_str("</fills>");

        xml.push_str("<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>");
        xml.push_str("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

        xml.push_str(&format!("<cellXfs count=\"{}\">", self.xfs.len()));
        for xf in &self.xfs {
            xml.push_str(&format!(
                "<xf numFmtId=\"{}\" fontId=\"{}\" fillId=\"{}\" borderId=\"0\" xfId=\"0\"",
                xf.num_fmt_id, xf.font_id, xf.fill_id
            ));
            if xf.num_fmt_id != 0 {
                xml.push_str(" applyNumberFormat=\"1\"");
            }
            if xf.font_id != 0 {
                xml.push_str(" applyFont=\"1\"");
            }
            if xf.fill_id != 0 {
                xml.push_str(" applyFill=\"1\"");
            }
            xml.push_str("/>");
        }
        xml.push_str("</cellXfs>");
        xml.push_str("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
        xml.push_str("<dxfs count=\"0\"/>");
        xml.push_str("<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>");
        xml.push_str("</styleSheet>");
        xml
    }
}
