use std::fs;
use std::path::Path;
use visi_core::core::vba::{VbaModule, VbaModuleKind, VbaProject};
use visi_core::core::vba_xlsx::build_vba_project_bin;

fn main() {
    let out_dir = Path::new("fuzz/seeds/vba_import");
    fs::create_dir_all(out_dir).unwrap();

    let empty = VbaProject::new_empty();
    let bin = build_vba_project_bin(&empty).unwrap();
    fs::write(out_dir.join("seed_empty_project.bin"), &bin).unwrap();

    let mut with_module = VbaProject::new_empty();
    with_module.modules.push(VbaModule {
        name: "Module1".to_string(),
        kind: VbaModuleKind::Standard,
        source: "Attribute VB_Name = \"Module1\"\r\nSub Foo()\r\n    Dim x As Integer\r\n    x = 1\r\nEnd Sub\r\n".to_string(),
        bound_sheet_id: None,
        prefix_bytes: with_module.seed_prefix_bytes.clone(),
        module_cookie: with_module.seed_module_cookie,
        cached_compressed_source: None,
    });
    let bin = build_vba_project_bin(&with_module).unwrap();
    fs::write(out_dir.join("seed_one_module.bin"), &bin).unwrap();

    println!("wrote seeds to {}", out_dir.display());
}
