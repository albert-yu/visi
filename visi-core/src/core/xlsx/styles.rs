use std::collections::HashMap;

use quick_xml::events::Event;

use crate::core::style::{
    BorderSide, BorderStyle, CellAlignment, CellBorders, CellStyle, FillPattern,
    HorizontalAlignment, VerticalAlignment,
};

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

#[derive(Debug, Clone)]
pub(crate) struct ThemeColors {
    palette: [String; 12],
}

impl Default for ThemeColors {
    fn default() -> Self {
        Self {
            palette: [
                "FFFFFF".to_string(),
                "000000".to_string(),
                "E7E6E6".to_string(),
                "44546A".to_string(),
                "4472C4".to_string(),
                "ED7D31".to_string(),
                "A5A5A5".to_string(),
                "FFC000".to_string(),
                "5B9BD5".to_string(),
                "70AD47".to_string(),
                "0563C1".to_string(),
                "954F72".to_string(),
            ],
        }
    }
}

impl ThemeColors {
    pub(crate) fn parse_theme_xml(xml: &str) -> Self {
        let mut theme = Self::default();
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buf = Vec::new();
        let mut in_clr_scheme = false;
        let mut active_slot: Option<usize> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"clrScheme" => in_clr_scheme = true,
                    local if in_clr_scheme && active_slot.is_none() => {
                        active_slot = theme_slot_index(local);
                    }
                    b"srgbClr" if active_slot.is_some() => {
                        if let Some(val) = super::get_attr(e, b"val")
                            && val.len() == 6
                            && let Some(slot) = active_slot
                        {
                            theme.palette[slot] = val.to_ascii_uppercase();
                        }
                    }
                    b"sysClr" if active_slot.is_some() => {
                        if let Some(val) = super::get_attr(e, b"lastClr")
                            && val.len() == 6
                            && let Some(slot) = active_slot
                        {
                            theme.palette[slot] = val.to_ascii_uppercase();
                        }
                    }
                    _ => {}
                },
                Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                    b"srgbClr" if active_slot.is_some() => {
                        if let Some(val) = super::get_attr(e, b"val")
                            && val.len() == 6
                            && let Some(slot) = active_slot
                        {
                            theme.palette[slot] = val.to_ascii_uppercase();
                        }
                    }
                    b"sysClr" if active_slot.is_some() => {
                        if let Some(val) = super::get_attr(e, b"lastClr")
                            && val.len() == 6
                            && let Some(slot) = active_slot
                        {
                            theme.palette[slot] = val.to_ascii_uppercase();
                        }
                    }
                    _ => {}
                },
                Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                    b"clrScheme" => {
                        in_clr_scheme = false;
                        active_slot = None;
                    }
                    local if in_clr_scheme && theme_slot_index(local) == active_slot => {
                        active_slot = None;
                    }
                    _ => {}
                },
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
        theme
    }

    pub(crate) fn resolve(&self, idx: usize, tint: Option<f64>) -> Option<String> {
        let base_hex = self.palette.get(idx)?;
        let r = u8::from_str_radix(&base_hex[0..2], 16).ok()?;
        let g = u8::from_str_radix(&base_hex[2..4], 16).ok()?;
        let b = u8::from_str_radix(&base_hex[4..6], 16).ok()?;
        let (r, g, b) = match tint {
            Some(t) if t.is_finite() && t != 0.0 => apply_tint(r, g, b, t.clamp(-1.0, 1.0)),
            _ => (r, g, b),
        };
        Some(format!("#{r:02X}{g:02X}{b:02X}"))
    }
}

fn theme_slot_index(local: &[u8]) -> Option<usize> {
    match local {
        b"lt1" => Some(0),
        b"dk1" => Some(1),
        b"lt2" => Some(2),
        b"dk2" => Some(3),
        b"accent1" => Some(4),
        b"accent2" => Some(5),
        b"accent3" => Some(6),
        b"accent4" => Some(7),
        b"accent5" => Some(8),
        b"accent6" => Some(9),
        b"hlink" => Some(10),
        b"folHlink" => Some(11),
        _ => None,
    }
}

fn apply_tint(r: u8, g: u8, b: u8, tint: f64) -> (u8, u8, u8) {
    let rf = r as f64 / 255.0;
    let gf = g as f64 / 255.0;
    let bf = b as f64 / 255.0;
    let max = rf.max(gf).max(bf);
    let min = rf.min(gf).min(bf);
    let delta = max - min;
    let mut l = (max + min) / 2.0;
    let s = if delta == 0.0 {
        0.0
    } else if l < 0.5 {
        delta / (max + min)
    } else {
        delta / (2.0 - max - min)
    };
    let h = if delta == 0.0 {
        0.0
    } else if max == rf {
        ((gf - bf) / delta + if gf < bf { 6.0 } else { 0.0 }) / 6.0
    } else if max == gf {
        ((bf - rf) / delta + 2.0) / 6.0
    } else {
        ((rf - gf) / delta + 4.0) / 6.0
    };

    l = if tint < 0.0 {
        l * (1.0 + tint)
    } else {
        l * (1.0 - tint) + tint
    }
    .clamp(0.0, 1.0);

    let (nr, ng, nb) = if s == 0.0 {
        (l, l, l)
    } else {
        let q = if l < 0.5 {
            l * (1.0 + s)
        } else {
            l + s - l * s
        };
        let p = 2.0 * l - q;
        (
            hue_to_rgb(p, q, h + 1.0 / 3.0),
            hue_to_rgb(p, q, h),
            hue_to_rgb(p, q, h - 1.0 / 3.0),
        )
    };

    (
        (nr * 255.0).round().clamp(0.0, 255.0) as u8,
        (ng * 255.0).round().clamp(0.0, 255.0) as u8,
        (nb * 255.0).round().clamp(0.0, 255.0) as u8,
    )
}

fn hue_to_rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 {
        t += 1.0;
    }
    if t > 1.0 {
        t -= 1.0;
    }
    if t < 1.0 / 6.0 {
        return p + (q - p) * 6.0 * t;
    }
    if t < 0.5 {
        return q;
    }
    if t < 2.0 / 3.0 {
        return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
    }
    p
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

#[derive(Debug, Clone, Default)]
struct ParsedFill {
    pattern: Option<FillPattern>,
    color: Option<String>,
}

#[derive(Debug, Clone, Copy)]
enum BorderSideTag {
    Left,
    Right,
    Top,
    Bottom,
    Diagonal,
}

pub(crate) fn parse_styles(xml: &str) -> ParsedStyles {
    parse_styles_with_theme(xml, &ThemeColors::default())
}

pub(crate) fn parse_styles_with_theme(xml: &str, theme: &ThemeColors) -> ParsedStyles {
    let mut custom_num_formats: HashMap<u32, String> = HashMap::new();
    let mut fonts: Vec<CellStyle> = Vec::new();
    let mut fills: Vec<ParsedFill> = Vec::new();
    let mut borders: Vec<CellBorders> = Vec::new();
    let mut xf_styles: HashMap<u32, CellStyle> = HashMap::new();
    let mut xf_is_date: Vec<bool> = Vec::new();

    let mut in_fonts = false;
    let mut in_font = false;
    let mut current_font = CellStyle::default();

    let mut in_fills = false;
    let mut in_fill = false;
    let mut current_fill = ParsedFill::default();

    let mut in_borders = false;
    let mut in_border = false;
    let mut current_border = CellBorders::default();
    let mut active_border_side: Option<(BorderSideTag, BorderStyle)> = None;

    let mut in_cell_xfs = false;
    let mut in_xf = false;
    let mut current_xf_style = CellStyle::default();
    let mut current_xf_is_date = false;
    let mut xf_idx = 0u32;

    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
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
                b"strike" if in_font => {
                    if style_bool_enabled(e) {
                        current_font.strikethrough = Some(true);
                    }
                }
                b"color" if in_font => current_font.font_color = style_color_hex(e, theme),
                b"name" if in_font => current_font.font_family = super::get_attr(e, b"val"),
                b"sz" if in_font => {
                    current_font.font_size =
                        super::get_attr(e, b"val").and_then(|s| s.parse().ok());
                }
                b"fills" => in_fills = true,
                b"fill" if in_fills => {
                    in_fill = true;
                    current_fill = ParsedFill::default();
                }
                b"patternFill" if in_fill => {
                    current_fill.pattern = super::get_attr(e, b"patternType")
                        .as_deref()
                        .and_then(FillPattern::from_ooxml);
                }
                b"fgColor" if in_fill => {
                    if let Some(color) = style_color_hex(e, theme) {
                        current_fill.color = Some(color);
                    }
                }
                b"bgColor" if in_fill => {
                    if current_fill.color.is_none() {
                        current_fill.color = style_color_hex(e, theme);
                    }
                }
                b"borders" => in_borders = true,
                b"border" if in_borders => {
                    in_border = true;
                    current_border = CellBorders {
                        diagonal_up: super::get_attr(e, b"diagonalUp")
                            .is_some_and(|v| v == "1" || v.eq_ignore_ascii_case("true")),
                        diagonal_down: super::get_attr(e, b"diagonalDown")
                            .is_some_and(|v| v == "1" || v.eq_ignore_ascii_case("true")),
                        ..Default::default()
                    };
                }
                b"left" | b"right" | b"top" | b"bottom" | b"diagonal" if in_border => {
                    let tag = match e.local_name().as_ref() {
                        b"left" => BorderSideTag::Left,
                        b"right" => BorderSideTag::Right,
                        b"top" => BorderSideTag::Top,
                        b"bottom" => BorderSideTag::Bottom,
                        _ => BorderSideTag::Diagonal,
                    };
                    if let Some(bs) = super::get_attr(e, b"style")
                        .as_deref()
                        .and_then(BorderStyle::from_ooxml)
                        && bs != BorderStyle::None
                    {
                        active_border_side = Some((tag, bs));
                        set_border_side(
                            &mut current_border,
                            tag,
                            BorderSide {
                                style: bs,
                                color: None,
                            },
                        );
                    } else {
                        active_border_side = None;
                    }
                }
                b"color" if active_border_side.is_some() => {
                    if let Some((tag, bs)) = active_border_side
                        && let Some(c) = style_color_hex(e, theme)
                    {
                        set_border_side(
                            &mut current_border,
                            tag,
                            BorderSide {
                                style: bs,
                                color: Some(c),
                            },
                        );
                    }
                }
                b"cellXfs" => in_cell_xfs = true,
                b"xf" if in_cell_xfs => {
                    in_xf = true;
                    let (style, is_date) =
                        build_xf_style(e, &fonts, &fills, &borders, &custom_num_formats);
                    current_xf_style = style;
                    current_xf_is_date = is_date;
                }
                b"alignment" if in_xf => {
                    if let Some(align) = parse_alignment_element(e) {
                        current_xf_style.alignment = Some(align);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"numFmt" => {
                    if let Some(id) =
                        super::get_attr(e, b"numFmtId").and_then(|s| s.parse::<u32>().ok())
                        && let Some(code) = super::get_attr(e, b"formatCode")
                    {
                        custom_num_formats.insert(id, code);
                    }
                }
                b"font" if in_fonts => fonts.push(CellStyle::default()),
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
                b"strike" if in_font => {
                    if style_bool_enabled(e) {
                        current_font.strikethrough = Some(true);
                    }
                }
                b"color" if in_font => current_font.font_color = style_color_hex(e, theme),
                b"name" if in_font => current_font.font_family = super::get_attr(e, b"val"),
                b"sz" if in_font => {
                    current_font.font_size =
                        super::get_attr(e, b"val").and_then(|s| s.parse().ok());
                }
                b"fill" if in_fills => fills.push(ParsedFill::default()),
                b"patternFill" if in_fill => {
                    current_fill.pattern = super::get_attr(e, b"patternType")
                        .as_deref()
                        .and_then(FillPattern::from_ooxml);
                }
                b"fgColor" if in_fill => {
                    if let Some(color) = style_color_hex(e, theme) {
                        current_fill.color = Some(color);
                    }
                }
                b"bgColor" if in_fill => {
                    if current_fill.color.is_none() {
                        current_fill.color = style_color_hex(e, theme);
                    }
                }
                b"border" if in_borders => borders.push(CellBorders::default()),
                b"left" | b"right" | b"top" | b"bottom" | b"diagonal" if in_border => {
                    let tag = match e.local_name().as_ref() {
                        b"left" => BorderSideTag::Left,
                        b"right" => BorderSideTag::Right,
                        b"top" => BorderSideTag::Top,
                        b"bottom" => BorderSideTag::Bottom,
                        _ => BorderSideTag::Diagonal,
                    };
                    if let Some(bs) = super::get_attr(e, b"style")
                        .as_deref()
                        .and_then(BorderStyle::from_ooxml)
                        && bs != BorderStyle::None
                    {
                        set_border_side(
                            &mut current_border,
                            tag,
                            BorderSide {
                                style: bs,
                                color: None,
                            },
                        );
                    }
                }
                b"color" if active_border_side.is_some() => {
                    if let Some((tag, bs)) = active_border_side
                        && let Some(c) = style_color_hex(e, theme)
                    {
                        set_border_side(
                            &mut current_border,
                            tag,
                            BorderSide {
                                style: bs,
                                color: Some(c),
                            },
                        );
                    }
                }
                b"xf" if in_cell_xfs => {
                    let (style, is_date) =
                        build_xf_style(e, &fonts, &fills, &borders, &custom_num_formats);
                    if !style.is_empty() {
                        xf_styles.insert(xf_idx, style);
                    }
                    xf_is_date.push(is_date);
                    xf_idx += 1;
                }
                b"alignment" if in_xf => {
                    if let Some(align) = parse_alignment_element(e) {
                        current_xf_style.alignment = Some(align);
                    }
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
                    current_fill = ParsedFill::default();
                    in_fill = false;
                }
                b"borders" => in_borders = false,
                b"border" if in_border => {
                    borders.push(current_border.clone());
                    current_border = CellBorders::default();
                    active_border_side = None;
                    in_border = false;
                }
                b"left" | b"right" | b"top" | b"bottom" | b"diagonal" if in_border => {
                    active_border_side = None;
                }
                b"xf" if in_xf => {
                    if !current_xf_style.is_empty() {
                        xf_styles.insert(xf_idx, current_xf_style.clone());
                    }
                    xf_is_date.push(current_xf_is_date);
                    xf_idx += 1;
                    current_xf_style = CellStyle::default();
                    current_xf_is_date = false;
                    in_xf = false;
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

fn set_border_side(border: &mut CellBorders, tag: BorderSideTag, side: BorderSide) {
    match tag {
        BorderSideTag::Left => border.left = Some(side),
        BorderSideTag::Right => border.right = Some(side),
        BorderSideTag::Top => border.top = Some(side),
        BorderSideTag::Bottom => border.bottom = Some(side),
        BorderSideTag::Diagonal => border.diagonal = Some(side),
    }
}

fn build_xf_style(
    e: &quick_xml::events::BytesStart,
    fonts: &[CellStyle],
    fills: &[ParsedFill],
    borders: &[CellBorders],
    custom_num_formats: &HashMap<u32, String>,
) -> (CellStyle, bool) {
    let mut style = CellStyle::default();
    if let Some(font_id) = super::get_attr(e, b"fontId").and_then(|s| s.parse::<usize>().ok())
        && font_id != 0
        && let Some(font) = fonts.get(font_id)
    {
        style.merge(font);
    }
    if let Some(fill_id) = super::get_attr(e, b"fillId").and_then(|s| s.parse::<usize>().ok())
        && fill_id > 1
        && let Some(fill) = fills.get(fill_id)
    {
        if let Some(ref bg_color) = fill.color {
            style.bg_color = Some(bg_color.clone());
        }
        if let Some(pat) = fill.pattern
            && pat != FillPattern::None
            && (pat != FillPattern::Solid || fill.color.is_none())
        {
            style.fill_pattern = Some(pat);
        }
    }
    if let Some(border_id) = super::get_attr(e, b"borderId").and_then(|s| s.parse::<usize>().ok())
        && border_id != 0
        && let Some(border) = borders.get(border_id)
        && !border.is_empty()
    {
        style.borders = Some(border.clone());
    }
    let mut is_date = false;
    if let Some(num_fmt_id) = super::get_attr(e, b"numFmtId").and_then(|s| s.parse::<u32>().ok()) {
        let code = custom_num_formats
            .get(&num_fmt_id)
            .cloned()
            .or_else(|| builtin_num_fmt_code(num_fmt_id).map(str::to_string));
        if let Some(ref c) = code {
            is_date = matches!(num_fmt_id, 14..=22 | 45..=47) || crate::core::date::is_date_code(c);
        }
        style.num_format = code.filter(|c| !c.eq_ignore_ascii_case("General"));
    }
    (style, is_date)
}

fn parse_alignment_element(e: &quick_xml::events::BytesStart) -> Option<CellAlignment> {
    let align = CellAlignment {
        horizontal: super::get_attr(e, b"horizontal")
            .as_deref()
            .and_then(HorizontalAlignment::from_ooxml),
        vertical: super::get_attr(e, b"vertical")
            .as_deref()
            .and_then(VerticalAlignment::from_ooxml),
        wrap_text: super::get_attr(e, b"wrapText")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true")),
        text_rotation: super::get_attr(e, b"textRotation").and_then(|s| s.parse::<u32>().ok()),
        indent: super::get_attr(e, b"indent").and_then(|s| s.parse::<u32>().ok()),
        shrink_to_fit: super::get_attr(e, b"shrinkToFit")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true")),
    };
    if align.is_empty() { None } else { Some(align) }
}

fn style_bool_enabled(e: &quick_xml::events::BytesStart) -> bool {
    super::get_attr(e, b"val").is_none_or(|value| {
        value != "0" && !value.eq_ignore_ascii_case("false") && !value.eq_ignore_ascii_case("none")
    })
}

fn style_color_hex(e: &quick_xml::events::BytesStart, theme: &ThemeColors) -> Option<String> {
    if let Some(rgb) = super::get_attr(e, b"rgb") {
        let trimmed = rgb.trim().trim_start_matches('#');
        let hex = match trimmed.len() {
            8 => &trimmed[2..],
            6 => trimmed,
            _ => return None,
        };
        return Some(format!("#{}", hex.to_ascii_uppercase()));
    }
    if let Some(theme_idx) = super::get_attr(e, b"theme").and_then(|s| s.parse::<usize>().ok()) {
        let tint = super::get_attr(e, b"tint").and_then(|s| s.parse::<f64>().ok());
        return theme.resolve(theme_idx, tint);
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
        16 => "#800000",
        17 => "#008000",
        18 => "#000080",
        19 => "#808000",
        20 => "#800080",
        21 => "#008080",
        22 => "#C0C0C0",
        23 => "#808080",
        24 => "#9999FF",
        25 => "#993366",
        26 => "#FFFFCC",
        27 => "#CCFFFF",
        28 => "#660066",
        29 => "#FF8080",
        30 => "#0066CC",
        31 => "#CCCCFF",
        32 => "#000080",
        33 => "#FF00FF",
        34 => "#FFFF00",
        35 => "#00FFFF",
        36 => "#800080",
        37 => "#800000",
        38 => "#008080",
        39 => "#0000FF",
        40 => "#00CCFF",
        41 => "#CCFFFF",
        42 => "#CCFFCC",
        43 => "#FFFF99",
        44 => "#99CCFF",
        45 => "#FF99CC",
        46 => "#CC99FF",
        47 => "#FFCC99",
        48 => "#3366FF",
        49 => "#33CCCC",
        50 => "#99CC00",
        51 => "#FFCC00",
        52 => "#FF9900",
        53 => "#FF6600",
        54 => "#666699",
        55 => "#969696",
        56 => "#003366",
        57 => "#339966",
        58 => "#003300",
        59 => "#333300",
        60 => "#993300",
        61 => "#993366",
        62 => "#333399",
        63 => "#333333",
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
    strikethrough: bool,
    size_tenths: Option<u32>,
    color_argb: Option<String>,
    family: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct FillKey {
    pattern: FillPattern,
    fg_argb: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct XfKey {
    num_fmt_id: u32,
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    alignment: Option<CellAlignment>,
}

#[derive(Debug)]
pub(crate) struct StyleRegistry {
    custom_num_fmts: Vec<(u32, String)>,
    num_fmt_map: HashMap<String, u32>,
    next_num_fmt_id: u32,
    fonts: Vec<FontKey>,
    font_map: HashMap<FontKey, usize>,
    fills: Vec<FillKey>,
    fill_map: HashMap<FillKey, usize>,
    borders: Vec<CellBorders>,
    border_map: HashMap<CellBorders, usize>,
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
            strikethrough: false,
            size_tenths: Some(110),
            color_argb: None,
            family: Some("Calibri".to_string()),
        };
        let mut font_map = HashMap::new();
        font_map.insert(default_font.clone(), 0);

        let default_fill_0 = FillKey {
            pattern: FillPattern::None,
            fg_argb: None,
        };
        let default_fill_1 = FillKey {
            pattern: FillPattern::Gray125,
            fg_argb: None,
        };

        let default_border = CellBorders::default();
        let mut border_map = HashMap::new();
        border_map.insert(default_border.clone(), 0);

        let default_xf = XfKey {
            num_fmt_id: 0,
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            alignment: None,
        };
        let mut xf_map = HashMap::new();
        xf_map.insert(default_xf.clone(), 0);

        Self {
            custom_num_fmts: Vec::new(),
            num_fmt_map,
            next_num_fmt_id: 164,
            fonts: vec![default_font],
            font_map,
            fills: vec![default_fill_0, default_fill_1],
            fill_map: HashMap::new(),
            borders: vec![default_border],
            border_map,
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
            || style.strikethrough.unwrap_or(false)
            || style.font_color.is_some()
            || style.font_family.is_some()
            || style.font_size.is_some();

        let font_id = if has_font_override {
            let key = FontKey {
                bold: style.bold.unwrap_or(false),
                italic: style.italic.unwrap_or(false),
                underline: style.underline.unwrap_or(false),
                strikethrough: style.strikethrough.unwrap_or(false),
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

        let fg_argb = style.bg_color.as_deref().and_then(parse_color_argb);
        let fill_id =
            if fg_argb.is_some() || style.fill_pattern.is_some_and(|p| p != FillPattern::None) {
                let pattern = style.fill_pattern.unwrap_or(FillPattern::Solid);
                let key = FillKey { pattern, fg_argb };
                if let Some(&id) = self.fill_map.get(&key) {
                    id
                } else {
                    let id = self.fills.len();
                    self.fills.push(key.clone());
                    self.fill_map.insert(key, id);
                    id
                }
            } else {
                0
            };

        let border_id = match style.borders.as_ref().filter(|b| !b.is_empty()) {
            Some(borders) => {
                if let Some(&id) = self.border_map.get(borders) {
                    id
                } else {
                    let id = self.borders.len();
                    self.borders.push(borders.clone());
                    self.border_map.insert(borders.clone(), id);
                    id
                }
            }
            None => 0,
        };

        let alignment = style.alignment.clone().filter(|a| !a.is_empty());

        let xf_key = XfKey {
            num_fmt_id,
            font_id,
            fill_id,
            border_id,
            alignment,
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
            if font.strikethrough {
                xml.push_str("<strike/>");
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
        for fill in &self.fills {
            let pat = fill.pattern.as_ooxml();
            if let Some(ref argb) = fill.fg_argb {
                xml.push_str(&format!(
                    "<fill><patternFill patternType=\"{pat}\"><fgColor rgb=\"{argb}\"/><bgColor indexed=\"64\"/></patternFill></fill>"
                ));
            } else {
                xml.push_str(&format!(
                    "<fill><patternFill patternType=\"{pat}\"/></fill>"
                ));
            }
        }
        xml.push_str("</fills>");

        xml.push_str(&format!("<borders count=\"{}\">", self.borders.len()));
        for border in &self.borders {
            xml.push_str("<border");
            if border.diagonal_up {
                xml.push_str(" diagonalUp=\"1\"");
            }
            if border.diagonal_down {
                xml.push_str(" diagonalDown=\"1\"");
            }
            xml.push('>');
            write_border_side(&mut xml, "left", border.left.as_ref());
            write_border_side(&mut xml, "right", border.right.as_ref());
            write_border_side(&mut xml, "top", border.top.as_ref());
            write_border_side(&mut xml, "bottom", border.bottom.as_ref());
            write_border_side(&mut xml, "diagonal", border.diagonal.as_ref());
            xml.push_str("</border>");
        }
        xml.push_str("</borders>");

        xml.push_str("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

        xml.push_str(&format!("<cellXfs count=\"{}\">", self.xfs.len()));
        for xf in &self.xfs {
            xml.push_str(&format!(
                "<xf numFmtId=\"{}\" fontId=\"{}\" fillId=\"{}\" borderId=\"{}\" xfId=\"0\"",
                xf.num_fmt_id, xf.font_id, xf.fill_id, xf.border_id
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
            if xf.border_id != 0 {
                xml.push_str(" applyBorder=\"1\"");
            }
            if let Some(ref align) = xf.alignment {
                xml.push_str(" applyAlignment=\"1\">");
                write_alignment_xml(&mut xml, align);
                xml.push_str("</xf>");
            } else {
                xml.push_str("/>");
            }
        }
        xml.push_str("</cellXfs>");
        xml.push_str("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
        xml.push_str("<dxfs count=\"0\"/>");
        xml.push_str("<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>");
        xml.push_str("</styleSheet>");
        xml
    }
}

fn write_border_side(xml: &mut String, tag: &str, side: Option<&BorderSide>) {
    match side {
        Some(s) if s.style != BorderStyle::None => {
            xml.push_str(&format!("<{tag} style=\"{}\"", s.style.as_ooxml()));
            if let Some(argb) = s.color.as_deref().and_then(parse_color_argb) {
                xml.push_str(&format!("><color rgb=\"{argb}\"/></{tag}>"));
            } else {
                xml.push_str("/>");
            }
        }
        _ => xml.push_str(&format!("<{tag}/>")),
    }
}

fn write_alignment_xml(xml: &mut String, align: &CellAlignment) {
    xml.push_str("<alignment");
    if let Some(h) = align.horizontal {
        xml.push_str(&format!(" horizontal=\"{}\"", h.as_ooxml()));
    }
    if let Some(v) = align.vertical {
        xml.push_str(&format!(" vertical=\"{}\"", v.as_ooxml()));
    }
    if let Some(rot) = align.text_rotation {
        xml.push_str(&format!(" textRotation=\"{rot}\""));
    }
    if let Some(true) = align.wrap_text {
        xml.push_str(" wrapText=\"1\"");
    }
    if let Some(indent) = align.indent {
        xml.push_str(&format!(" indent=\"{indent}\""));
    }
    if let Some(true) = align.shrink_to_fit {
        xml.push_str(" shrinkToFit=\"1\"");
    }
    xml.push_str("/>");
}
