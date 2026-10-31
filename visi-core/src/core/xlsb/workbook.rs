use super::record::{write_empty_record, write_record, write_wide_string};

fn write_file_version(out: &mut Vec<u8>) {
    let mut payload = vec![0; 16];
    write_wide_string("xl", &mut payload);
    write_wide_string("7", &mut payload);
    write_wide_string("7", &mut payload);
    write_wide_string("29822", &mut payload);
    write_record(0x0080, &payload, out);
}

pub(crate) fn build(sheet_names: &[String]) -> Vec<u8> {
    let mut out = Vec::with_capacity(sheet_names.len() * 64 + 192);
    write_empty_record(0x0083, &mut out);
    write_file_version(&mut out);
    write_record(
        0x0099,
        &[
            0x20, 0x00, 0x01, 0x00, 0x42, 0xe5, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
        ],
        &mut out,
    );
    write_empty_record(0x0087, &mut out);
    write_record(0x0025, &[0x01, 0x00, 0x02, 0x10, 0x00, 0x80], &mut out);
    write_record(
        0x0c00,
        &[
            0x00, 0x00, 0x00, 0x00, 0x0d, 0x00, 0x00, 0x00, 0xff, 0xff, 0xff, 0xff, 0x00, 0x00,
            0x00, 0x00,
        ],
        &mut out,
    );
    write_empty_record(0x0026, &mut out);
    write_record(
        0x009e,
        &[
            0x89, 0x1c, 0x00, 0x00, 0x56, 0x13, 0x00, 0x00, 0xbf, 0x5e, 0x00, 0x00, 0x5b, 0x3b,
            0x00, 0x00, 0x58, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x78,
        ],
        &mut out,
    );
    write_empty_record(0x0088, &mut out);
    write_empty_record(0x008f, &mut out);
    for (index, name) in sheet_names.iter().enumerate() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&((index + 1) as u32).to_le_bytes());
        write_wide_string(&format!("rId{}", index + 1), &mut payload);
        write_wide_string(name, &mut payload);
        write_record(0x009c, &payload, &mut out);
    }
    write_empty_record(0x0090, &mut out);
    write_empty_record(0x0161, &mut out);
    write_empty_record(0x0165, &mut out);
    write_record(
        0x016a,
        &[
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00,
        ],
        &mut out,
    );
    write_empty_record(0x0162, &mut out);
    write_record(
        0x009d,
        &[
            0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00, 0xfc, 0xa9,
            0xf1, 0xd2, 0x4d, 0x62, 0x50, 0x3f, 0x01, 0x00, 0x00, 0x00, 0x6a, 0x00,
        ],
        &mut out,
    );
    write_record(0x009b, &[0x00], &mut out);
    write_empty_record(0x0084, &mut out);
    out
}
