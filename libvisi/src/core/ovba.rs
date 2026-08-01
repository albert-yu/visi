//! MS-OVBA "Compressed Container" codec (MS-OVBA section 2.4), used for both
//! a `vbaProject.bin`'s `dir` stream and every module's source-code stream.
//!
//! Decompression is a direct implementation of the documented format.
//! Compression is a real greedy LZ77 (hash-indexed match finder) rather than
//! a naive literal-only or "stored" encoder: real Excel-authored data was
//! found (empirically, via a scratchpad proof-of-concept validated against
//! real Excel) to always shrink each non-final chunk's 4096 decompressed
//! bytes into <=4096 encoded bytes using genuine back-references, and to
//! never use the spec-legal "stored/uncompressed" flag=0 chunk type -- both
//! a naive literal encoder (which can't fit 4096 literal bytes in a
//! 4096-byte budget) and a stored-chunk workaround were confirmed to make
//! Excel silently drop the affected module from the VBA project tree, even
//! though both decompress correctly through this same decompressor.

pub fn decompress(data: &[u8]) -> Vec<u8> {
    assert_eq!(data[0], 0x01, "bad compressed container signature");
    let mut out = Vec::new();
    let mut pos = 1;
    while pos < data.len() {
        let header = u16::from_le_bytes([data[pos], data[pos + 1]]);
        let chunk_size = (header & 0x0FFF) as usize + 3; // includes the 2-byte header
        let signature = (header >> 12) & 0b111;
        assert_eq!(signature, 0b011, "bad chunk signature at pos {pos}");
        let compressed_flag = (header >> 15) & 1;
        let chunk_start = pos + 2;
        let chunk_data_end = pos + chunk_size; // exclusive, relative to `data`
        let chunk_data = &data[chunk_start..chunk_data_end.min(data.len())];

        if compressed_flag == 0 {
            // Uncompressed chunk: exactly 4096 raw bytes. Not produced by
            // `compress` below, but real files could in principle contain
            // one, so decoding still supports it.
            out.extend_from_slice(chunk_data);
        } else {
            decompress_chunk(chunk_data, &mut out);
        }
        pos = chunk_data_end;
    }
    out
}

fn decompress_chunk(chunk_data: &[u8], out: &mut Vec<u8>) {
    let chunk_start_out = out.len();
    let mut i = 0;
    while i < chunk_data.len() {
        let flag_byte = chunk_data[i];
        i += 1;
        for bit in 0..8 {
            if i >= chunk_data.len() {
                break;
            }
            let is_copy_token = (flag_byte >> bit) & 1 == 1;
            if !is_copy_token {
                out.push(chunk_data[i]);
                i += 1;
            } else {
                let token = u16::from_le_bytes([chunk_data[i], chunk_data[i + 1]]);
                i += 2;
                let decompressed_current = out.len() - chunk_start_out;
                let bit_count = bit_count_for(decompressed_current);
                let length_mask: u16 = 0xFFFF >> bit_count;
                let offset_mask: u16 = !length_mask;
                let length = (token & length_mask) as usize + 3;
                let offset = (((token & offset_mask) >> (16 - bit_count)) + 1) as usize;
                let copy_start = out.len() - offset;
                for k in 0..length {
                    let b = out[copy_start + k];
                    out.push(b);
                }
            }
        }
    }
}

/// MS-OVBA 2.4.1.3.19: smallest number of bits (clamped 4..=12) such that
/// `decompressed_current <= 2^bit_count` -- governs the offset/length bit
/// split for a copy token at this point in the chunk.
fn bit_count_for(decompressed_current: usize) -> u32 {
    let mut bit_count = 0u32;
    while (1usize << bit_count) < decompressed_current {
        bit_count += 1;
    }
    bit_count.clamp(4, 12)
}

/// Compresses `data` into an MS-OVBA Compressed Container using real LZ77
/// back-references, chunked to exactly 4096 decompressed bytes per
/// non-final chunk (matching genuine Excel-authored data).
pub fn compress(data: &[u8]) -> Vec<u8> {
    let mut out = vec![0x01u8];
    let mut remaining = data;
    while !remaining.is_empty() {
        let take = remaining.len().min(4096);
        let (chunk, rest) = remaining.split_at(take);
        let body = compress_chunk(chunk);
        let total_size = 2 + body.len();
        assert!(
            total_size <= 4098,
            "chunk body too large to encode: {} bytes",
            body.len()
        );
        let chunk_size_field = (total_size - 3) as u16;
        let header: u16 = (1 << 15) | (0b011 << 12) | chunk_size_field;
        out.extend_from_slice(&header.to_le_bytes());
        out.extend_from_slice(&body);
        remaining = rest;
    }
    out
}

fn compress_chunk(input: &[u8]) -> Vec<u8> {
    use std::collections::HashMap;
    let mut body = Vec::new();
    let mut hash_index: HashMap<[u8; 3], Vec<usize>> = HashMap::new();
    let mut i = 0;
    while i < input.len() {
        let mut flag_byte = 0u8;
        let mut group_body = Vec::new();
        for bit in 0..8 {
            if i >= input.len() {
                break;
            }
            let decompressed_current = i;
            let bit_count = bit_count_for(decompressed_current);
            let length_mask: u16 = 0xFFFF >> bit_count;
            let max_length = length_mask as usize + 3;
            let max_offset = 1usize << bit_count;

            let best = if i + 3 <= input.len() {
                find_best_match(input, i, max_offset, max_length, &hash_index)
            } else {
                None
            };

            if let Some((offset, length)) = best {
                let token: u16 =
                    (((offset - 1) as u16) << (16 - bit_count)) | ((length - 3) as u16);
                group_body.extend_from_slice(&token.to_le_bytes());
                flag_byte |= 1 << bit;
                for k in i..(i + length).min(input.len()) {
                    if k + 3 <= input.len() {
                        let key = [input[k], input[k + 1], input[k + 2]];
                        let entries = hash_index.entry(key).or_default();
                        entries.push(k);
                        if entries.len() > 64 {
                            entries.remove(0);
                        }
                    }
                }
                i += length;
            } else {
                group_body.push(input[i]);
                if i + 3 <= input.len() {
                    let key = [input[i], input[i + 1], input[i + 2]];
                    let entries = hash_index.entry(key).or_default();
                    entries.push(i);
                    if entries.len() > 64 {
                        entries.remove(0);
                    }
                }
                i += 1;
            }
        }
        body.push(flag_byte);
        body.extend_from_slice(&group_body);
    }
    body
}

fn find_best_match(
    input: &[u8],
    i: usize,
    max_offset: usize,
    max_length: usize,
    hash_index: &std::collections::HashMap<[u8; 3], Vec<usize>>,
) -> Option<(usize, usize)> {
    let key = [input[i], input[i + 1], input[i + 2]];
    let candidates = hash_index.get(&key)?;
    let min_j = i.saturating_sub(max_offset);
    let mut best_offset = 0usize;
    let mut best_length = 0usize;
    for &j in candidates.iter().rev() {
        if j < min_j || j >= i {
            continue;
        }
        let mut length = 0;
        while i + length < input.len()
            && length < max_length
            && input[j + length] == input[i + length]
        {
            length += 1;
        }
        if length > best_length {
            best_length = length;
            best_offset = i - j;
        }
    }
    if best_length >= 3 {
        Some((best_offset, best_length))
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn roundtrip_multichunk_boundaries() {
        for len in [1, 3640, 3641, 4095, 4096, 4097, 8192, 8193, 12000] {
            let original: Vec<u8> = (0..len).map(|i| (i % 251) as u8).collect();
            let compressed = compress(&original);
            let decompressed = decompress(&compressed);
            assert_eq!(decompressed, original, "length {len} failed");
        }
    }

    #[test]
    fn roundtrip_repetitive_text() {
        let original = b"Sub Test()\n    Dim x As Integer\n    x = 1\nEnd Sub\n".repeat(200);
        let compressed = compress(&original);
        let decompressed = decompress(&compressed);
        assert_eq!(decompressed, original);
        assert!(compressed.len() < original.len());
    }

    #[test]
    fn roundtrip_empty() {
        let compressed = compress(b"");
        let decompressed = decompress(&compressed);
        assert_eq!(decompressed, b"");
    }
}
