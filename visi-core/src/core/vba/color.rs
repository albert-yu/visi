use crate::core::CellStyle;

pub(crate) const NO_FILL_COLOR: i32 = 16_777_215;

pub(crate) const COLOR_INDEX_NONE: i32 = -4142;

pub(crate) const FONT_COLOR_INDEX_AUTOMATIC: i32 = 1;

pub(crate) const DEFAULT_FONT_SIZE: f64 = 11.0;

pub(crate) const DEFAULT_FONT_NAME: &str = "Calibri";

pub(crate) const GENERAL_FORMAT: &str = "General";

pub(crate) const COLOR_INDEX_PALETTE: [&str; 56] = [
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
    "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
    "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
    "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
    "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
    "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
    "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];

pub(crate) fn rgb(r: i64, g: i64, b: i64) -> i32 {
    let clamp = |v: i64| v.clamp(0, 255) as i32;
    clamp(r) | (clamp(g) << 8) | (clamp(b) << 16)
}

pub(crate) fn bgr_to_hex(color: i32) -> String {
    let v = color as u32;
    let (r, g, b) = (v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF);
    format!("#{r:02X}{g:02X}{b:02X}")
}

pub(crate) fn hex_to_bgr(text: &str) -> Option<i32> {
    let hex = named_color(text).unwrap_or_else(|| text.trim().trim_start_matches('#'));
    if hex.len() != 6 || !hex.chars().all(|c| c.is_ascii_hexdigit()) {
        return None;
    }
    let r = i32::from_str_radix(&hex[0..2], 16).ok()?;
    let g = i32::from_str_radix(&hex[2..4], 16).ok()?;
    let b = i32::from_str_radix(&hex[4..6], 16).ok()?;
    Some(r | (g << 8) | (b << 16))
}

pub(crate) fn nearest_color_index(hex: &str) -> Option<i32> {
    let (r, g, b) = components(hex)?;
    let mut best = (u32::MAX, 0usize);
    for (i, slot) in COLOR_INDEX_PALETTE.iter().enumerate() {
        let (pr, pg, pb) = components(slot).expect("palette entries are valid hex");
        let d = |a: i32, b: i32| ((a - b) * (a - b)) as u32;
        let distance = d(r, pr) + d(g, pg) + d(b, pb);
        if distance < best.0 {
            best = (distance, i);
        }
    }
    Some(best.1 as i32 + 1)
}

fn components(text: &str) -> Option<(i32, i32, i32)> {
    let bgr = hex_to_bgr(text)?;
    Some((bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF))
}

fn named_color(text: &str) -> Option<&'static str> {
    Some(match text.trim().to_ascii_lowercase().as_str() {
        "black" => "000000",
        "white" => "FFFFFF",
        "red" => "FF0000",
        "green" => "00FF00",
        "blue" => "0000FF",
        "yellow" => "FFFF00",
        "magenta" => "FF00FF",
        "cyan" => "00FFFF",
        "gray" | "grey" => "808080",
        _ => return None,
    })
}

pub(crate) fn interior_color(style: Option<&CellStyle>) -> i32 {
    style
        .and_then(|s| s.bg_color.as_deref())
        .and_then(hex_to_bgr)
        .unwrap_or(NO_FILL_COLOR)
}

pub(crate) fn font_color(style: Option<&CellStyle>) -> i32 {
    style
        .and_then(|s| s.font_color.as_deref())
        .and_then(hex_to_bgr)
        .unwrap_or(0)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rgb_packs_the_components_the_way_excel_does() {
        assert_eq!(rgb(255, 0, 0), 255);
        assert_eq!(rgb(0, 255, 0), 65280);
        assert_eq!(rgb(0, 0, 255), 16_711_680);
        assert_eq!(rgb(1, 2, 3), 197_121);
        assert_eq!(rgb(300, 0, 0), 255);
    }

    #[test]
    fn the_long_is_bgr_so_ff0000_is_blue() {
        assert_eq!(bgr_to_hex(0x00FF_0000), "#0000FF");
        assert_eq!(bgr_to_hex(255), "#FF0000");
        assert_eq!(bgr_to_hex(rgb(1, 2, 3)), "#010203");
        assert_eq!(hex_to_bgr("#0000FF"), Some(16_711_680));
        assert_eq!(hex_to_bgr("#FF0000"), Some(255));
    }

    #[test]
    fn a_colour_round_trips_through_both_representations() {
        for color in [0, 255, 65280, 16_711_680, 197_121, 16_777_215] {
            assert_eq!(hex_to_bgr(&bgr_to_hex(color)), Some(color), "{color}");
        }
    }

    #[test]
    fn colour_names_and_bare_hex_are_read_back_too() {
        assert_eq!(hex_to_bgr("red"), Some(255));
        assert_eq!(hex_to_bgr("FF0000"), Some(255));
        assert_eq!(hex_to_bgr("not a colour"), None);
    }

    #[test]
    fn the_palette_is_the_one_excel_reported() {
        assert_eq!(nearest_color_index("#FF0000"), Some(3));
        assert_eq!(nearest_color_index("#000000"), Some(1));
        assert_eq!(nearest_color_index("#FFFFFF"), Some(2));
        assert_eq!(nearest_color_index("#333333"), Some(56));
        assert_eq!(nearest_color_index("#000080"), Some(11));
    }

    #[test]
    fn an_off_palette_colour_reports_the_nearest_slot_not_xlnone() {
        assert_eq!(nearest_color_index("#010203"), Some(1), "near black");
        assert_eq!(nearest_color_index("#FA0A0A"), Some(3), "near red");
        assert_eq!(nearest_color_index("#0AC80A"), Some(4), "near green");
    }

    #[test]
    fn every_palette_slot_survives_the_colour_round_trip() {
        for (i, hex) in COLOR_INDEX_PALETTE.iter().enumerate() {
            let bgr = hex_to_bgr(hex).unwrap_or_else(|| panic!("slot {}", i + 1));
            assert_eq!(&bgr_to_hex(bgr), hex, "slot {}", i + 1);
        }
    }

    #[test]
    fn an_unstyled_cell_reports_excels_defaults() {
        assert_eq!(interior_color(None), 16_777_215);
        assert_eq!(font_color(None), 0);
    }
}
