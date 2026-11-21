#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if let Ok(src) = std::str::from_utf8(data)
        && let Ok(module) = visi_core::core::vba::parser::parse_module(src)
        && let Some(first) = module.procedures().first()
    {
        let proc_name = first.name.clone();
        let mut interp = visi_core::core::vba::interp::Interpreter::new(module).with_max_ops(50_000);
        let _ = interp.run(&proc_name, Vec::new());
    }
});
