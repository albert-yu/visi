//! VBA's colour representation, and the bridge to `CellStyle`'s.
//!
//! `Interior.Color` and `Font.Color` are a `Long` in **BGR** order -- the
//! low byte is red -- while [`CellStyle`](crate::core::CellStyle) stores
//! `"#RRGGBB"`. So `&HFF0000` is **blue**, not red, and an implementation
//! that treats the `Long` as `0xRRGGBB` produces a file that opens without
//! complaint and is the wrong colour. Issue #58 called this out as the single
//! most likely thing to ship backwards, so it is measured rather than
//! recalled: `fuzz/vba_style_probe.py --paint` has Excel *save* a workbook
//! after setting colours and reads the real ARGB back with `openpyxl`, which
//! is a channel neither implementation can talk its way past.
//!
//! Measured, and what [`rgb`] and [`bgr_to_hex`] encode:
//!
//! | VBA | in the saved file |
//! | --- | --- |
//! | `RGB(255, 0, 0)` = `255` | `FFFF0000` (red) |
//! | `&HFF0000` = `16711680` | `FF0000FF` (blue) |
//! | `RGB(1, 2, 3)` = `197121` | `FF010203` |

use crate::core::CellStyle;

/// `Interior.Color` on a cell with no fill. Measured: white, not zero and not
/// an error.
pub(crate) const NO_FILL_COLOR: i32 = 16_777_215;

/// `Interior.ColorIndex` on a cell with no fill -- `xlNone`. Measured.
pub(crate) const COLOR_INDEX_NONE: i32 = -4142;

/// `Font.ColorIndex` on a cell with no explicit font colour. Measured:
/// slot 1, which is black, rather than `xlNone`.
pub(crate) const FONT_COLOR_INDEX_AUTOMATIC: i32 = 1;

/// `Font.Size` on an unstyled cell, and the size a new cell is rendered at.
/// Measured; note Excel reports it as a `Double`, not a `Long`.
pub(crate) const DEFAULT_FONT_SIZE: f64 = 11.0;

/// `Font.Name` on an unstyled cell. Measured.
pub(crate) const DEFAULT_FONT_NAME: &str = "Calibri";

/// `Range.NumberFormat` on a cell carrying no format. Measured.
pub(crate) const GENERAL_FORMAT: &str = "General";

/// Excel's 56-slot `ColorIndex` palette, slot 1 first.
///
/// Every entry was read out of Excel by `fuzz/vba_style_probe.py --palette`,
/// which sets `Interior.ColorIndex = n` and reads the resulting `Color` back,
/// rather than transcribed from a published table. Note that the palette is
/// not injective -- slots 25..32 repeat 9..16 (`#000080` is both 11 and 25),
/// which is why [`nearest_color_index`] breaks ties toward the lower slot and why
/// `Color` is the representation `CellStyle` keeps.
pub(crate) const COLOR_INDEX_PALETTE: [&str; 56] = [
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
    "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
    "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
    "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
    "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
    "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
    "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];

/// `RGB(r, g, b)` -- the `Long` VBA composes from three components.
///
/// Measured: components above 255 clamp rather than overflowing into the next
/// byte (`RGB(300, 0, 0)` is 255), and a negative component is error 5, which
/// is the caller's job to raise since this returns the value only.
pub(crate) fn rgb(r: i64, g: i64, b: i64) -> i32 {
    let clamp = |v: i64| v.clamp(0, 255) as i32;
    clamp(r) | (clamp(g) << 8) | (clamp(b) << 16)
}

/// A VBA colour `Long` as the `"#RRGGBB"` [`CellStyle`] stores.
///
/// The byte swap is the whole point; see this module's doc comment.
pub(crate) fn bgr_to_hex(color: i32) -> String {
    let v = color as u32;
    let (r, g, b) = (v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF);
    format!("#{r:02X}{g:02X}{b:02X}")
}

/// A `CellStyle` colour string as the `Long` VBA reports.
///
/// Accepts what `CellStyle` accepts: `"#RRGGBB"`, a bare `"RRGGBB"`, and the
/// handful of colour names the CLI's `--color` takes, so a fill set by
/// `visi style` reads back through a macro rather than erroring.
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

/// The palette slot Excel reports for a colour.
///
/// **Not** an exact-match lookup: measured, Excel maps an off-palette colour
/// to the *nearest* slot rather than reporting `xlNone`. `RGB(250, 10, 10)`
/// reports slot 3 (pure red), `RGB(10, 200, 10)` reports slot 4 (pure green),
/// and `RGB(1, 2, 3)` reports slot 1 (black). An exact match falls out of the
/// same search at distance zero.
///
/// Ties go to the lower slot, which is what makes an exact `#000080` report
/// 11 rather than 25 -- the palette repeats itself from slot 25 on.
///
/// The metric is squared Euclidean distance in RGB, which reproduces every
/// measured point; it is a fit to three off-palette samples rather than a
/// documented rule, so a case that disagrees is a reason to re-measure rather
/// than to assume this is right.
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

/// The `(r, g, b)` components of a colour string, in 0..=255.
fn components(text: &str) -> Option<(i32, i32, i32)> {
    let bgr = hex_to_bgr(text)?;
    Some((bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF))
}

/// The colour names `CellStyle` documents, so a style set through the CLI is
/// legible to a macro. Not Excel's -- Excel has no colour names in this
/// position at all; this is only about reading back what visi itself wrote.
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

/// What `Interior.Color` reports for a cell, styled or not.
pub(crate) fn interior_color(style: Option<&CellStyle>) -> i32 {
    style
        .and_then(|s| s.bg_color.as_deref())
        .and_then(hex_to_bgr)
        .unwrap_or(NO_FILL_COLOR)
}

/// What `Font.Color` reports for a cell. Measured: an unstyled cell is 0,
/// i.e. black, rather than an error or a sentinel.
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
        // Each measured through `CStr(RGB(...))`.
        assert_eq!(rgb(255, 0, 0), 255);
        assert_eq!(rgb(0, 255, 0), 65280);
        assert_eq!(rgb(0, 0, 255), 16_711_680);
        assert_eq!(rgb(1, 2, 3), 197_121);
        // Measured: a component over 255 clamps rather than carrying into the
        // next byte, which would have turned `RGB(300, 0, 0)` green.
        assert_eq!(rgb(300, 0, 0), 255);
    }

    #[test]
    fn the_long_is_bgr_so_ff0000_is_blue() {
        // The case issue #58 named as the most likely thing to ship
        // backwards, and the one `--paint` settles against the saved file.
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
        // Not Excel behaviour -- this is so a fill set by `visi style
        // --bg-color red` is legible to a macro rather than erroring.
        assert_eq!(hex_to_bgr("red"), Some(255));
        assert_eq!(hex_to_bgr("FF0000"), Some(255));
        assert_eq!(hex_to_bgr("not a colour"), None);
    }

    #[test]
    fn the_palette_is_the_one_excel_reported() {
        // Spot checks against `--palette`; the whole table came from that run.
        assert_eq!(nearest_color_index("#FF0000"), Some(3));
        assert_eq!(nearest_color_index("#000000"), Some(1));
        assert_eq!(nearest_color_index("#FFFFFF"), Some(2));
        assert_eq!(nearest_color_index("#333333"), Some(56));
        // The palette repeats: `#000080` is slots 11 and 25, and Excel
        // reports the lower one.
        assert_eq!(nearest_color_index("#000080"), Some(11));
    }

    #[test]
    fn an_off_palette_colour_reports_the_nearest_slot_not_xlnone() {
        // All three measured. The first guess here was `xlNone`, and it was
        // wrong -- Excel does a nearest-colour match.
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
