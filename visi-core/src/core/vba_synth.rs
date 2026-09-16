use crate::core::ovba;
use crate::core::vba_xlsx::write_record;
use std::io::Write;

fn write_project_version_record(out: &mut Vec<u8>) {
    out.extend_from_slice(&0x0009u16.to_le_bytes());
    out.extend_from_slice(&4u32.to_le_bytes());
    out.extend_from_slice(&1u32.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
}

fn build_skeleton_dir() -> Vec<u8> {
    let mut dir = Vec::new();
    write_record(&mut dir, 0x0001, &1u32.to_le_bytes());
    write_record(&mut dir, 0x004A, &0x0002_0000u32.to_le_bytes());
    write_record(&mut dir, 0x0002, &0x0409u32.to_le_bytes());
    write_record(&mut dir, 0x0014, &0x0409u32.to_le_bytes());
    write_record(&mut dir, 0x0003, &0x2710u16.to_le_bytes());
    write_record(&mut dir, 0x0004, b"VBAProject");
    write_record(&mut dir, 0x0005, &[]);
    write_record(&mut dir, 0x0040, &[]);
    write_record(&mut dir, 0x0006, &[]);
    write_record(&mut dir, 0x003D, &[]);
    write_record(&mut dir, 0x0007, &0u32.to_le_bytes());
    write_record(&mut dir, 0x0008, &0u32.to_le_bytes());
    write_project_version_record(&mut dir);
    write_record(&mut dir, 0x000C, &[]);
    write_record(&mut dir, 0x003C, &[]);
    write_record(&mut dir, 0x000F, &0u16.to_le_bytes());
    write_record(&mut dir, 0x0013, &0xFFFFu16.to_le_bytes());
    write_record(&mut dir, 0x0010, &[]);
    dir
}

fn build_skeleton_vba_project_cache() -> Vec<u8> {
    let mut cache = Vec::new();
    cache.extend_from_slice(&0x61CCu16.to_le_bytes());
    cache.extend_from_slice(&0x00DFu16.to_le_bytes());
    cache.push(0x00);
    cache.extend_from_slice(&[0x00, 0x00]);
    cache
}

/// Builds a from-scratch `vbaProject.bin` CFB container holding just the
/// `dir` and `_VBA_PROJECT` streams -- everything `VbaProject::raw_donor`
/// needs to exist for `vba_xlsx::build_vba_project_bin` to patch, without
/// any of it being copied from a real file.
pub fn synthetic_raw_donor() -> Vec<u8> {
    let dir_compressed =
        ovba::compress(&build_skeleton_dir()).expect("tiny synthetic dir stream always compresses");
    let vba_project_cache = build_skeleton_vba_project_cache();

    let mut cf =
        cfb::CompoundFile::create_with_version(cfb::Version::V3, std::io::Cursor::new(Vec::new()))
            .expect("in-memory CFB container creation cannot fail");
    cf.create_storage("VBA")
        .expect("fresh CFB storage creation cannot fail");
    cf.create_stream("VBA/dir")
        .and_then(|mut s| s.write_all(&dir_compressed))
        .expect("fresh CFB stream write cannot fail");
    cf.create_stream("VBA/_VBA_PROJECT")
        .and_then(|mut s| s.write_all(&vba_project_cache))
        .expect("fresh CFB stream write cannot fail");
    cf.into_inner().into_inner()
}

const OBJECT_TABLE_BASE_OFFSET: usize = 0x05;
const INDIRECT_TABLE_BASE_OFFSET: usize = 0x11;
const LINE_TABLE_BASE_OFFSET: usize = 0x19;
const DECL_TABLE_LEN_OFFSET: usize = 0x3F;

const INDIRECT_TABLE_LEN_OFFSET: usize = DECL_TABLE_LEN_OFFSET + 4;
const OBJECT_TABLE_LEN_OFFSET: usize = 0x8A;
const MAGIC_OFFSET: usize = OBJECT_TABLE_LEN_OFFSET + 4;
const LINE_COUNT_OFFSET: usize = MAGIC_OFFSET + 4;
const PREFIX_LEN: usize = LINE_COUNT_OFFSET + 2;

const CAFE_MAGIC: u16 = 0xCAFE;

/// A minimal, self-consistent, zero-procedure p-code prefix -- see the
/// module doc comment for why this replaces borrowing real bytes from a
/// donor module.
pub fn synthetic_module_prefix() -> Vec<u8> {
    let mut buf = vec![0u8; PREFIX_LEN];

    buf[OBJECT_TABLE_BASE_OFFSET..OBJECT_TABLE_BASE_OFFSET + 4]
        .copy_from_slice(&0u32.to_le_bytes());

    let indirect_table_base = (INDIRECT_TABLE_LEN_OFFSET - 10) as u32;
    buf[INDIRECT_TABLE_BASE_OFFSET..INDIRECT_TABLE_BASE_OFFSET + 4]
        .copy_from_slice(&indirect_table_base.to_le_bytes());

    let line_table_base = (MAGIC_OFFSET - 0x3C) as u32;
    buf[LINE_TABLE_BASE_OFFSET..LINE_TABLE_BASE_OFFSET + 4]
        .copy_from_slice(&line_table_base.to_le_bytes());

    buf[MAGIC_OFFSET..MAGIC_OFFSET + 2].copy_from_slice(&CAFE_MAGIC.to_le_bytes());

    buf
}

#[cfg(test)]
mod tests {
    use super::*;

    fn read_u32(buf: &[u8], offset: usize) -> u32 {
        u32::from_le_bytes(buf[offset..offset + 4].try_into().unwrap())
    }

    fn read_u16(buf: &[u8], offset: usize) -> u16 {
        u16::from_le_bytes(buf[offset..offset + 2].try_into().unwrap())
    }

    #[test]
    fn synthetic_prefix_is_self_consistent() {
        let buf = synthetic_module_prefix();

        let decl_len = read_u32(&buf, DECL_TABLE_LEN_OFFSET);
        assert_eq!(decl_len, 0, "declaration table must be empty");

        let indirect_base = read_u32(&buf, INDIRECT_TABLE_BASE_OFFSET) as usize;
        let indirect_len_offset = indirect_base + 10;
        assert_eq!(
            read_u32(&buf, indirect_len_offset),
            0,
            "indirect table must be empty"
        );

        let object_base = read_u32(&buf, OBJECT_TABLE_BASE_OFFSET) as usize;
        let object_len_offset = object_base + 0x8A;
        assert_eq!(
            read_u32(&buf, object_len_offset),
            0,
            "object table must be empty"
        );

        let line_table_base = read_u32(&buf, LINE_TABLE_BASE_OFFSET) as usize;
        let magic_offset = line_table_base + 0x3C;
        assert_eq!(read_u16(&buf, magic_offset), CAFE_MAGIC);
        let line_count_offset = magic_offset + 2 + 2;
        assert_eq!(
            read_u16(&buf, line_count_offset),
            0,
            "line count must be zero"
        );

        assert!(line_count_offset + 2 <= buf.len());
        assert!(indirect_len_offset + 4 <= buf.len());
        assert!(object_len_offset + 4 <= buf.len());
    }

    #[test]
    fn synthetic_raw_donor_has_expected_streams() {
        let bytes = synthetic_raw_donor();
        let mut cfb_file = cfb::CompoundFile::open(std::io::Cursor::new(bytes)).unwrap();
        assert!(cfb_file.open_stream("/VBA/dir").is_ok());
        assert!(cfb_file.open_stream("/VBA/_VBA_PROJECT").is_ok());
    }
}
