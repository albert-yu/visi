//! Synthesizes a brand-new VBA project's binary internals from scratch: the
//! `vbaProject.bin` "donor" skeleton (`dir` + `_VBA_PROJECT` streams) that
//! `vba_xlsx::build_vba_project_bin` patches on every export, plus each
//! module's opaque p-code prefix. Together these let a workbook's first VBA
//! module be created without copying any bytes from a real, Excel-authored
//! file -- see `vba.rs` for why that used to be necessary.
//!
//! The `dir`-stream record IDs used by `build_skeleton_dir` (PROJECTSYSKIND,
//! PROJECTLCID, PROJECTVERSION, ...) are documented in the public
//! [MS-OVBA] specification. The module p-code prefix's internal layout is
//! not: MS-OVBA calls it a "PerformanceCache" that implementations "MAY
//! ignore", but real Excel was found (via the scratchpad proof-of-concept
//! referenced from `vba.rs`) to actually parse it and silently drop a
//! module whose bytes don't fit the expected shape, rather than falling
//! back to the module's source text. `synthetic_module_prefix` builds a
//! self-consistent, *zero-line* cache -- empty declaration/indirect/object
//! tables, the `0xCAFE` line-table marker, a line count of zero -- at the
//! fixed absolute offsets that shape is understood to occupy, so there is
//! nothing for Excel to execute out of the cache and it always (re)compiles
//! the module fresh from source instead. That makes this one fixed byte
//! sequence valid for every module, independent of its actual source text.

use crate::core::ovba;
use crate::core::vba_xlsx::write_record;
use std::io::Write;

/// [MS-OVBA] 2.3.4.1 PROJECTVERSION: a fixed 12-byte record regardless of
/// its nominal size field (`vba_xlsx::read_dir_record` special-cases it).
fn write_project_version_record(out: &mut Vec<u8>) {
    out.extend_from_slice(&0x0009u16.to_le_bytes());
    out.extend_from_slice(&4u32.to_le_bytes()); // Reserved, MUST be 0x00000004
    out.extend_from_slice(&1u32.to_le_bytes()); // VersionMajor
    out.extend_from_slice(&0u16.to_le_bytes()); // VersionMinor
}

/// A complete, reference-free `dir` stream: PROJECTINFORMATION with an
/// empty PROJECTREFERENCES section, then an empty PROJECTMODULES list.
/// `vba_xlsx::build_vba_project_bin` only ever reads the PROJECTINFORMATION
/// prefix (everything before PROJECTMODULES) out of `raw_donor`'s dir
/// stream and rebuilds PROJECTMODULES itself from the live `VbaProject`, so
/// the empty module list written here is never actually read back -- it
/// exists only so this dir stream is valid and inspectable on its own.
fn build_skeleton_dir() -> Vec<u8> {
    let mut dir = Vec::new();
    write_record(&mut dir, 0x0001, &1u32.to_le_bytes()); // PROJECTSYSKIND: Win32
    write_record(&mut dir, 0x004A, &0x0002_0000u32.to_le_bytes()); // PROJECTCOMPATVERSION
    write_record(&mut dir, 0x0002, &0x0409u32.to_le_bytes()); // PROJECTLCID
    write_record(&mut dir, 0x0014, &0x0409u32.to_le_bytes()); // PROJECTLCIDINVOKE
    write_record(&mut dir, 0x0003, &0x2710u16.to_le_bytes()); // PROJECTCODEPAGE
    write_record(&mut dir, 0x0004, b"VBAProject"); // PROJECTNAME
    write_record(&mut dir, 0x0005, &[]); // PROJECTDOCSTRING
    write_record(&mut dir, 0x0040, &[]); // PROJECTDOCSTRING (Unicode)
    write_record(&mut dir, 0x0006, &[]); // PROJECTHELPFILEPATH1
    write_record(&mut dir, 0x003D, &[]); // PROJECTHELPFILEPATH2
    write_record(&mut dir, 0x0007, &0u32.to_le_bytes()); // PROJECTHELPCONTEXT
    write_record(&mut dir, 0x0008, &0u32.to_le_bytes()); // PROJECTLIBFLAGS
    write_project_version_record(&mut dir); // PROJECTVERSION
    write_record(&mut dir, 0x000C, &[]); // PROJECTCONSTANTS
    write_record(&mut dir, 0x003C, &[]); // PROJECTCONSTANTS (Unicode)
    write_record(&mut dir, 0x000F, &0u16.to_le_bytes()); // PROJECTMODULES, count=0
    write_record(&mut dir, 0x0013, &0xFFFFu16.to_le_bytes()); // PROJECTCOOKIE
    write_record(&mut dir, 0x0010, &[]); // PROJECTTERMINATOR
    dir
}

/// The whole-project `_VBA_PROJECT` cache stream: just the 7-byte header
/// ([MS-OVBA] 2.3.4.3 -- Reserved1 = 0x61CC, an implementation-defined
/// version tag, Reserved2 = 0x00, Reserved3), with no cached data trailing
/// it.
fn build_skeleton_vba_project_cache() -> Vec<u8> {
    let mut cache = Vec::new();
    cache.extend_from_slice(&0x61CCu16.to_le_bytes()); // Reserved1
    cache.extend_from_slice(&0x00DFu16.to_le_bytes()); // Version
    cache.push(0x00); // Reserved2
    cache.extend_from_slice(&[0x00, 0x00]); // Reserved3
    cache
}

/// Builds a from-scratch `vbaProject.bin` CFB container holding just the
/// `dir` and `_VBA_PROJECT` streams -- everything `VbaProject::raw_donor`
/// needs to exist for `vba_xlsx::build_vba_project_bin` to patch, without
/// any of it being copied from a real file.
pub fn synthetic_raw_donor() -> Vec<u8> {
    let dir_compressed = ovba::compress(&build_skeleton_dir());
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

// ---------------------------------------------------------------------------
// Module p-code prefix
// ---------------------------------------------------------------------------

// Absolute byte offsets a module stream's p-code prefix is understood to be
// read at (see the module doc comment). Each "base" offset holds a dword
// that, plus a fixed constant, locates a table-length dword elsewhere in the
// buffer; picking each base so its table-length dword lands immediately
// after the previous field keeps the whole buffer minimal.
const OBJECT_TABLE_BASE_OFFSET: usize = 0x05; // + 0x8A -> object table length
const INDIRECT_TABLE_BASE_OFFSET: usize = 0x11; // + 10 -> indirect table length
const LINE_TABLE_BASE_OFFSET: usize = 0x19; // + 0x3C -> magic + line count
const DECL_TABLE_LEN_OFFSET: usize = 0x3F; // declaration table length, left 0

const INDIRECT_TABLE_LEN_OFFSET: usize = DECL_TABLE_LEN_OFFSET + 4; // 0x43
const OBJECT_TABLE_LEN_OFFSET: usize = 0x8A; // OBJECT_TABLE_BASE_OFFSET's value (0) + 0x8A
const MAGIC_OFFSET: usize = OBJECT_TABLE_LEN_OFFSET + 4;
const LINE_COUNT_OFFSET: usize = MAGIC_OFFSET + 4; // magic word + 2 reserved bytes
const PREFIX_LEN: usize = LINE_COUNT_OFFSET + 2;

const CAFE_MAGIC: u16 = 0xCAFE;

/// A minimal, self-consistent, zero-procedure p-code prefix -- see the
/// module doc comment for why this replaces borrowing real bytes from a
/// donor module.
pub fn synthetic_module_prefix() -> Vec<u8> {
    let mut buf = vec![0u8; PREFIX_LEN];

    // Object table base stays 0 (buffer is already zero-filled), placing
    // its length dword at the fixed OBJECT_TABLE_LEN_OFFSET; written
    // explicitly so the relationship is visible next to its two siblings
    // below rather than left implicit.
    buf[OBJECT_TABLE_BASE_OFFSET..OBJECT_TABLE_BASE_OFFSET + 4]
        .copy_from_slice(&0u32.to_le_bytes());

    let indirect_table_base = (INDIRECT_TABLE_LEN_OFFSET - 10) as u32;
    buf[INDIRECT_TABLE_BASE_OFFSET..INDIRECT_TABLE_BASE_OFFSET + 4]
        .copy_from_slice(&indirect_table_base.to_le_bytes());

    let line_table_base = (MAGIC_OFFSET - 0x3C) as u32;
    buf[LINE_TABLE_BASE_OFFSET..LINE_TABLE_BASE_OFFSET + 4]
        .copy_from_slice(&line_table_base.to_le_bytes());

    buf[MAGIC_OFFSET..MAGIC_OFFSET + 2].copy_from_slice(&CAFE_MAGIC.to_le_bytes());

    // Declaration/indirect/object table lengths and the line count all stay
    // zero -- the buffer starts zero-filled, so nothing more to write.
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

    /// Independently walks `synthetic_module_prefix`'s output the same way
    /// a reader is understood to (see the module doc comment), as a
    /// regression guard on the layout rather than a proof it satisfies
    /// Excel.
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

        // Every field read above must fall within the buffer we actually
        // produced -- guards against a future offset change quietly
        // reading (or requiring) out-of-bounds bytes.
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
