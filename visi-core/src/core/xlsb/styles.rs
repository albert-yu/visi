use crate::core::{CellStyle, CellType, date::is_date_code};

use super::record::{write_empty_record, write_record};

pub(crate) fn build() -> Vec<u8> {
    let mut out = Vec::new();
    write_empty_record(0x0116, &mut out);
    write_record(0x0263, &1u32.to_le_bytes(), &mut out);
    write_record(
        0x002b,
        &[
            0xdc, 0x00, 0x00, 0x00, 0x90, 0x01, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x07, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0x02, 0x07, 0x00, 0x00, 0x00, 0x43, 0x00, 0x61,
            0x00, 0x6c, 0x00, 0x69, 0x00, 0x62, 0x00, 0x72, 0x00, 0x69, 0x00,
        ],
        &mut out,
    );
    write_empty_record(0x0264, &mut out);
    write_record(0x025b, &2u32.to_le_bytes(), &mut out);
    for pattern in [0x00, 0x11] {
        let mut payload = [0u8; 68];
        payload[0] = pattern;
        payload[4..16].copy_from_slice(&[
            0x03, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0x03, 0x41, 0x00, 0x00,
        ]);
        payload[16..20].copy_from_slice(&0xffff_ffffu32.to_le_bytes());
        write_record(0x002d, &payload, &mut out);
    }
    write_empty_record(0x025c, &mut out);
    write_record(0x0265, &1u32.to_le_bytes(), &mut out);
    write_record(
        0x002e,
        &[
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        ],
        &mut out,
    );
    write_empty_record(0x0266, &mut out);
    write_record(0x0272, &1u32.to_le_bytes(), &mut out);
    write_record(
        0x002f,
        &[
            0xff, 0xff, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10,
            0x00, 0x00,
        ],
        &mut out,
    );
    write_empty_record(0x0273, &mut out);
    write_record(0x0269, &2u32.to_le_bytes(), &mut out);
    write_record(
        0x002f,
        &[
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10,
            0x00, 0x00,
        ],
        &mut out,
    );
    write_record(
        0x002f,
        &[
            0x00, 0x00, 0x0e, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x08, 0x00,
            0x25, 0x00,
        ],
        &mut out,
    );
    write_empty_record(0x026a, &mut out);
    write_record(0x026b, &1u32.to_le_bytes(), &mut out);
    write_record(
        0x0030,
        &[
            0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, 0x4e, 0x00,
            0x6f, 0x00, 0x72, 0x00, 0x6d, 0x00, 0x61, 0x00, 0x6c, 0x00,
        ],
        &mut out,
    );
    write_empty_record(0x026c, &mut out);
    write_record(0x01f9, &0u32.to_le_bytes(), &mut out);
    write_empty_record(0x01fa, &mut out);
    write_empty_record(0x0117, &mut out);
    out
}

pub(crate) fn style_index(style: Option<&CellStyle>, cell_type: CellType) -> Result<u32, String> {
    let date_type = matches!(
        cell_type,
        CellType::DateTime | CellType::DateTimeIso | CellType::DurationIso
    );
    let Some(style) = style else {
        return Ok(u32::from(date_type));
    };
    let only_number_format = style.font_color.is_none()
        && style.bg_color.is_none()
        && style.bold.is_none()
        && style.italic.is_none()
        && style.underline.is_none()
        && style.font_family.is_none()
        && style.font_size.is_none();
    if style.is_empty()
        || only_number_format
            && style
                .num_format
                .as_deref()
                .is_some_and(|format| format.eq_ignore_ascii_case("General"))
    {
        Ok(u32::from(date_type))
    } else if only_number_format
        && (date_type || style.num_format.as_deref().is_some_and(is_date_code))
    {
        Ok(1)
    } else {
        Err(format!(
            "cell formatting other than date number formats is not supported in XLSB output: {style:?}"
        ))
    }
}
