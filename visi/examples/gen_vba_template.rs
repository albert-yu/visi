// One-off generator for `libvisi/assets/vba_template.bin`. Not part of the
// shipped CLI -- run manually (`cargo run --example gen_vba_template`)
// whenever the template needs regenerating. Borrows ONE real module's
// p-code prefix bytes from a local, real Excel-authored macro-enabled
// workbook (verified, via a scratchpad proof-of-concept against real Excel,
// that this prefix must be genuine p-code-shaped bytes -- zero-filled
// garbage of the same length makes Excel silently drop the module). Only
// the prefix bytes are borrowed; the module's actual source text is
// replaced with a generic placeholder.
use std::io::{Read, Write};

fn write_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn main() {
    let donor_xlsm_path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "fuzz/pivot_macro_template.xlsm".to_string());
    let donor_bytes = std::fs::read(&donor_xlsm_path)
        .unwrap_or_else(|e| panic!("failed to read {donor_xlsm_path}: {e}"));

    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(&donor_bytes[..]))
        .expect("donor file is not a valid zip/xlsm");
    let mut vba_bin = Vec::new();
    zip.by_name("xl/vbaProject.bin")
        .expect("donor file has no xl/vbaProject.bin")
        .read_to_end(&mut vba_bin)
        .unwrap();

    let mut donor_cfb =
        cfb::CompoundFile::open(std::io::Cursor::new(&vba_bin[..])).expect("bad CFB container");
    let mut dir_raw = Vec::new();
    donor_cfb
        .open_stream("VBA/dir")
        .unwrap()
        .read_to_end(&mut dir_raw)
        .unwrap();
    let dir = libvisi::core::ovba::decompress(&dir_raw);

    // Find Module1's MODULEOFFSET (0x0031) by walking the dir stream's
    // PROJECTMODULES section generically.
    let text_offset = find_module1_text_offset(&dir);
    let mut module1_raw = Vec::new();
    donor_cfb
        .open_stream("VBA/Module1")
        .unwrap()
        .read_to_end(&mut module1_raw)
        .unwrap();
    let real_prefix = module1_raw[..text_offset].to_vec();
    println!("borrowed {} real prefix bytes from donor Module1", real_prefix.len());

    // Build a minimal, zero-reference dir stream + minimal _VBA_PROJECT
    // cache skeleton -- proven sufficient (Test U/S in the proof-of-concept)
    // as long as each module's own prefix bytes are real.
    let mut new_dir = Vec::new();
    write_record(&mut new_dir, 0x0001, &3u32.to_le_bytes());
    write_record(&mut new_dir, 0x004A, &0x00020000u32.to_le_bytes());
    write_record(&mut new_dir, 0x0002, &0x0409u32.to_le_bytes());
    write_record(&mut new_dir, 0x0014, &0x0409u32.to_le_bytes());
    write_record(&mut new_dir, 0x0003, &0x2710u16.to_le_bytes());
    write_record(&mut new_dir, 0x0004, b"VBAProject");
    write_record(&mut new_dir, 0x0005, &[]);
    write_record(&mut new_dir, 0x0040, &[]);
    write_record(&mut new_dir, 0x0006, &[]);
    write_record(&mut new_dir, 0x003D, &[]);
    write_record(&mut new_dir, 0x0007, &0u32.to_le_bytes());
    write_record(&mut new_dir, 0x0008, &0u32.to_le_bytes());
    new_dir.extend_from_slice(&0x0009u16.to_le_bytes());
    new_dir.extend_from_slice(&[0u8; 4]);
    new_dir.extend_from_slice(&1u32.to_le_bytes());
    new_dir.extend_from_slice(&0u16.to_le_bytes());
    write_record(&mut new_dir, 0x000C, &[]);
    write_record(&mut new_dir, 0x003C, &[]);
    // Empty PROJECTMODULES section: export_vba_project only ever reads the
    // dir-stream PREFIX (everything before this record) from raw_donor, and
    // reads _VBA_PROJECT verbatim -- it never reads raw_donor's own module
    // list, so this can be a genuinely empty placeholder.
    write_record(&mut new_dir, 0x000F, &0u16.to_le_bytes());
    write_record(&mut new_dir, 0x0013, &0xFFFFu16.to_le_bytes());
    write_record(&mut new_dir, 0x0010, &[]);
    let new_dir_compressed = libvisi::core::ovba::compress(&new_dir);

    let vba_project_cache: Vec<u8> = {
        let mut c = Vec::new();
        c.extend_from_slice(&0x61CCu16.to_le_bytes());
        c.extend_from_slice(&0x00DFu16.to_le_bytes());
        c.push(0x00);
        c.extend_from_slice(&[0x00, 0x00]);
        c
    };

    let mut cf =
        cfb::CompoundFile::create_with_version(cfb::Version::V3, std::io::Cursor::new(Vec::new()))
            .unwrap();
    cf.create_storage("VBA").unwrap();
    cf.create_stream("VBA/dir")
        .unwrap()
        .write_all(&new_dir_compressed)
        .unwrap();
    cf.create_stream("VBA/_VBA_PROJECT")
        .unwrap()
        .write_all(&vba_project_cache)
        .unwrap();
    let skeleton_raw_donor = cf.into_inner().into_inner();

    // Now build the actual template VbaProject (one Standard module, real
    // borrowed prefix, generic placeholder source) and export it through
    // libvisi's own (just-written) export path, exactly the same code path
    // real usage will exercise.
    let template_project = libvisi::core::VbaProject {
        project_id: "{7B4E3A2C-1F5D-4A6B-9C8E-2D3F4A5B6C7D}".to_string(),
        modules: vec![libvisi::core::VbaModule {
            name: "Module1".to_string(),
            kind: libvisi::core::VbaModuleKind::Standard,
            source: "Attribute VB_Name = \"Module1\"\r\n\
' Placeholder module bundled with libvisi, replaced when a real module is added.\r\n\
Sub Placeholder()\r\nEnd Sub\r\n"
                .to_string(),
            bound_sheet_id: None,
            prefix_bytes: real_prefix,
        }],
        raw_donor: skeleton_raw_donor,
        seed_prefix_bytes: Vec::new(),
    };

    let template_bin = libvisi::core::vba_xlsx::build_vba_project_bin(&template_project)
        .expect("failed to build template vbaProject.bin");

    std::fs::write("libvisi/assets/vba_template.bin", &template_bin).unwrap();
    println!("wrote {} bytes to libvisi/assets/vba_template.bin", template_bin.len());
}

fn find_module1_text_offset(dir: &[u8]) -> usize {
    let mut pos = 0;
    while pos + 6 <= dir.len() {
        let id = u16::from_le_bytes([dir[pos], dir[pos + 1]]);
        if id == 0x0009 {
            pos += 12;
            continue;
        }
        let size = u32::from_le_bytes([dir[pos + 2], dir[pos + 3], dir[pos + 4], dir[pos + 5]]) as usize;
        let data_start = pos + 6;
        let data = &dir[data_start..data_start + size];
        if id == 0x0019 && data == b"Module1" {
            // Walk forward to this module's 0x0031 MODULEOFFSET record.
            let mut p = data_start + size;
            loop {
                let rid = u16::from_le_bytes([dir[p], dir[p + 1]]);
                let rsize = u32::from_le_bytes([dir[p + 2], dir[p + 3], dir[p + 4], dir[p + 5]]) as usize;
                let rdata = &dir[p + 6..p + 6 + rsize];
                if rid == 0x0031 {
                    return u32::from_le_bytes([rdata[0], rdata[1], rdata[2], rdata[3]]) as usize;
                }
                p += 6 + rsize;
            }
        }
        pos = data_start + size;
    }
    panic!("Module1 not found in donor dir stream");
}
