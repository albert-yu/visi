//! MS-OVBA "Compressed Container" codec (MS-OVBA section 2.4), used for both
//! a `vbaProject.bin`'s `dir` stream and every module's source-code stream.
//!
//! Decompression is a direct implementation of the documented format, fully
//! bounds-checked so malformed/truncated input (any `.xlsx`/`.xlsm` this
//! codebase didn't author itself) returns a `Result` rather than panicking.
//! Compression is a real greedy LZ77 (hash-indexed match finder) rather than
//! a naive literal-only or "stored" encoder: real Excel-authored data was
//! found (empirically, via a scratchpad proof-of-concept validated against
//! real Excel) to always shrink each non-final chunk's 4096 decompressed
//! bytes into <=4096 encoded bytes using genuine back-references, and to
//! never use the spec-legal "stored/uncompressed" flag=0 chunk type -- both
//! a naive literal encoder (which can't fit 4096 literal bytes in a
//! 4096-byte budget) and a stored-chunk workaround were confirmed to make
//! Excel silently drop the affected module from the VBA project tree, even
//! though both decompress correctly through this same decompressor. That
//! means a 4096-byte chunk with no 3-byte repeat anywhere in it (an
//! adversarial/high-entropy input, not realistic VBA source) genuinely
//! cannot be encoded in this format at all -- `compress` reports that as an
//! error rather than silently falling back to the one encoding already
//! proven to corrupt real Excel's macro loading.

pub fn decompress(data: &[u8]) -> Result<Vec<u8>, String> {
    let sig = *data
        .first()
        .ok_or("empty compressed container: missing signature byte")?;
    if sig != 0x01 {
        return Err(format!(
            "bad compressed container signature: expected 0x01, got 0x{sig:02X}"
        ));
    }
    let mut out = Vec::new();
    let mut pos = 1;
    while pos < data.len() {
        let header_bytes = data
            .get(pos..pos + 2)
            .ok_or("truncated compressed container: chunk header cut off")?;
        let header = u16::from_le_bytes([header_bytes[0], header_bytes[1]]);
        let chunk_size = (header & 0x0FFF) as usize + 3; // includes the 2-byte header
        let signature = (header >> 12) & 0b111;
        if signature != 0b011 {
            return Err(format!(
                "bad chunk signature 0b{signature:03b} at pos {pos}"
            ));
        }
        let compressed_flag = (header >> 15) & 1;
        let chunk_start = pos + 2;
        let chunk_data_end = pos + chunk_size; // exclusive, relative to `data`
        let chunk_data = data
            .get(chunk_start..chunk_data_end.min(data.len()))
            .ok_or("truncated compressed container: chunk body cut off")?;

        if compressed_flag == 0 {
            // Uncompressed chunk: exactly 4096 raw bytes. Not produced by
            // `compress` below, but real files could in principle contain
            // one, so decoding still supports it.
            out.extend_from_slice(chunk_data);
        } else {
            decompress_chunk(chunk_data, &mut out)?;
        }
        pos = chunk_data_end;
    }
    Ok(out)
}

fn decompress_chunk(chunk_data: &[u8], out: &mut Vec<u8>) -> Result<(), String> {
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
                let token_bytes = chunk_data
                    .get(i..i + 2)
                    .ok_or("truncated copy token in compressed chunk")?;
                let token = u16::from_le_bytes([token_bytes[0], token_bytes[1]]);
                i += 2;
                let decompressed_current = out.len() - chunk_start_out;
                let bit_count = bit_count_for(decompressed_current);
                let length_mask: u16 = 0xFFFF >> bit_count;
                let offset_mask: u16 = !length_mask;
                let length = (token & length_mask) as usize + 3;
                let offset = (((token & offset_mask) >> (16 - bit_count)) + 1) as usize;
                let copy_start = out
                    .len()
                    .checked_sub(offset)
                    .ok_or("copy token offset points before the start of the output")?;
                for k in 0..length {
                    let b = out[copy_start + k];
                    out.push(b);
                }
            }
        }
    }
    Ok(())
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
/// non-final chunk (matching genuine Excel-authored data). Errors (rather
/// than falling back to a "stored" chunk -- see the module doc comment for
/// why) if some 4096-byte chunk has so few repeated 3-byte sequences that
/// this encoding's per-byte overhead can't fit it in the format's budget.
pub fn compress(data: &[u8]) -> Result<Vec<u8>, String> {
    let mut out = vec![0x01u8];
    let mut remaining = data;
    while !remaining.is_empty() {
        let take = remaining.len().min(4096);
        let (chunk, rest) = remaining.split_at(take);
        let body = compress_chunk(chunk);
        let total_size = 2 + body.len();
        if total_size > 4098 {
            return Err(format!(
                "VBA source can't be encoded: a {}-byte block has too few repeated \
                 sequences to fit Excel's compressed module format (needs {} bytes, \
                 budget is 4096)",
                chunk.len(),
                body.len()
            ));
        }
        let chunk_size_field = (total_size - 3) as u16;
        let header: u16 = (1 << 15) | (0b011 << 12) | chunk_size_field;
        out.extend_from_slice(&header.to_le_bytes());
        out.extend_from_slice(&body);
        remaining = rest;
    }
    Ok(out)
}

fn compress_chunk(input: &[u8]) -> Vec<u8> {
    use std::collections::{HashMap, VecDeque};
    let mut body = Vec::new();
    let mut hash_index: HashMap<[u8; 3], VecDeque<usize>> = HashMap::new();
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
                        entries.push_back(k);
                        if entries.len() > 64 {
                            entries.pop_front();
                        }
                    }
                }
                i += length;
            } else {
                group_body.push(input[i]);
                if i + 3 <= input.len() {
                    let key = [input[i], input[i + 1], input[i + 2]];
                    let entries = hash_index.entry(key).or_default();
                    entries.push_back(i);
                    if entries.len() > 64 {
                        entries.pop_front();
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
    hash_index: &std::collections::HashMap<[u8; 3], std::collections::VecDeque<usize>>,
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
    use proptest::prelude::*;

    proptest! {
        // `compress` can legitimately reject high-entropy input (see the
        // module doc comment), so this only checks the roundtrip property
        // when compression succeeds -- `decompress_never_panics_or_ooms`
        // below covers arbitrary bytes on the decode side, which is the
        // side that actually receives untrusted input from imported files.
        #[test]
        fn roundtrip_when_compressible(data in proptest::collection::vec(any::<u8>(), 0..4096)) {
            if let Ok(compressed) = compress(&data) {
                let decompressed = decompress(&compressed).unwrap();
                prop_assert_eq!(decompressed, data);
            }
        }

        #[test]
        fn decompress_never_panics_or_ooms(data in proptest::collection::vec(any::<u8>(), 0..4096)) {
            let _ = decompress(&data);
        }
    }

    #[test]
    fn roundtrip_multichunk_boundaries() {
        for len in [1, 3640, 3641, 4095, 4096, 4097, 8192, 8193, 12000] {
            let original: Vec<u8> = (0..len).map(|i| (i % 251) as u8).collect();
            let compressed = compress(&original).unwrap();
            let decompressed = decompress(&compressed).unwrap();
            assert_eq!(decompressed, original, "length {len} failed");
        }
    }

    #[test]
    fn roundtrip_repetitive_text() {
        let original = b"Sub Test()\n    Dim x As Integer\n    x = 1\nEnd Sub\n".repeat(200);
        let compressed = compress(&original).unwrap();
        let decompressed = decompress(&compressed).unwrap();
        assert_eq!(decompressed, original);
        assert!(compressed.len() < original.len());
    }

    #[test]
    fn roundtrip_empty() {
        let compressed = compress(b"").unwrap();
        let decompressed = decompress(&compressed).unwrap();
        assert_eq!(decompressed, b"");
    }

    #[test]
    fn compress_errors_instead_of_panicking_on_incompressible_chunk() {
        // A 4096-byte chunk of xorshift32 pseudorandom bytes: high-entropy
        // enough that 3-byte-window repeats are rare (expected well under
        // one across 4096 positions among 256^3 possible 3-tuples), so this
        // encoding's literal-plus-flag-byte overhead (4096 literals need
        // 4096 + 512 flag bytes = 4608) can't fit the 4096-byte budget.
        let mut state = 0x2463_9A11u32;
        let original: Vec<u8> = (0..4096)
            .map(|_| {
                state ^= state << 13;
                state ^= state >> 17;
                state ^= state << 5;
                state as u8
            })
            .collect();
        assert!(compress(&original).is_err());
    }

    #[test]
    fn decompress_errors_instead_of_panicking_on_malformed_input() {
        assert!(decompress(&[]).is_err());
        assert!(decompress(&[0x02]).is_err()); // bad signature
        assert!(decompress(&[0x01, 0x00]).is_err()); // truncated header
        assert!(decompress(&[0x01, 0x00, 0x00]).is_err()); // bad chunk signature bits
        // Compressed chunk whose sole token claims an offset larger than
        // anything decoded so far (would underflow `out.len() - offset`).
        let mut malformed = vec![0x01u8];
        let body = [0x01u8, 0xFF, 0xFF]; // flag byte: bit 0 set (copy token) + 2 token bytes
        let total_size = 2 + body.len();
        let header: u16 = (1 << 15) | (0b011 << 12) | (total_size - 3) as u16;
        malformed.extend_from_slice(&header.to_le_bytes());
        malformed.extend_from_slice(&body);
        assert!(decompress(&malformed).is_err());
    }
}
