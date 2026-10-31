pub(crate) fn write_varint(mut value: u32, out: &mut Vec<u8>) {
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        if value == 0 {
            out.push(byte);
            break;
        }
        out.push(byte | 0x80);
    }
}

pub(crate) fn write_record(id: u32, payload: &[u8], out: &mut Vec<u8>) {
    write_varint(id, out);
    write_varint(payload.len() as u32, out);
    out.extend_from_slice(payload);
}

pub(crate) fn write_empty_record(id: u32, out: &mut Vec<u8>) {
    write_varint(id, out);
    out.push(0);
}

pub(crate) fn write_wide_string(value: &str, out: &mut Vec<u8>) {
    let encoded = value.encode_utf16().collect::<Vec<_>>();
    out.extend_from_slice(&(encoded.len() as u32).to_le_bytes());
    for unit in encoded {
        out.extend_from_slice(&unit.to_le_bytes());
    }
}

pub(crate) fn encode_rk(value: f64) -> Option<u32> {
    if !value.is_finite() {
        return None;
    }
    let integer = value as i64;
    if integer as f64 == value && (-(1 << 29)..(1 << 29)).contains(&integer) {
        return Some(((integer << 2) | 2) as u32);
    }
    let scaled = value * 100.0;
    let scaled_integer = scaled.round() as i64;
    if (scaled - scaled_integer as f64).abs() < 1e-6
        && (-(1 << 29)..(1 << 29)).contains(&scaled_integer)
    {
        return Some(((scaled_integer << 2) | 3) as u32);
    }
    let high = (value.to_bits() >> 32) as u32;
    if high & 3 == 0 { Some(high) } else { None }
}

pub(crate) fn cell_header(col: usize, style: u32) -> [u8; 8] {
    let mut header = [0u8; 8];
    header[..4].copy_from_slice(&(col as u32).to_le_bytes());
    let style_bytes = style.to_le_bytes();
    header[4..7].copy_from_slice(&style_bytes[..3]);
    header
}

pub(crate) fn write_row_header(
    row: usize,
    first_col: usize,
    last_col: usize,
    height: Option<f64>,
    out: &mut Vec<u8>,
) -> Result<(), String> {
    let mut payload = [0u8; 25];
    payload[..4].copy_from_slice(&(row as u32).to_le_bytes());
    let height = match height {
        Some(height) if height.is_finite() && height > 0.0 && height <= 409.5 => {
            payload[11] = 0x20;
            (height * 20.0).round() as u16
        }
        Some(_) => return Err("row heights must be finite and between 0 and 409.5".to_string()),
        None => 300,
    };
    payload[8..10].copy_from_slice(&height.to_le_bytes());
    payload[13..17].copy_from_slice(&1u32.to_le_bytes());
    payload[17..21].copy_from_slice(&(first_col as u32).to_le_bytes());
    payload[21..25].copy_from_slice(&(last_col as u32).to_le_bytes());
    write_record(0, &payload, out);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn record_header_uses_biff12_varints() {
        let mut out = Vec::new();
        write_record(0x9c, &[1, 2, 3], &mut out);
        assert_eq!(out, [0x9c, 0x01, 0x03, 1, 2, 3]);
    }

    #[test]
    fn rk_encodes_integers_and_hundredths() {
        assert_eq!(encode_rk(10.0), Some(42));
        assert_eq!(encode_rk(9.99), Some((999 << 2) | 3));
    }
}
