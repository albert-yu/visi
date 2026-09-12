# Differential Fuzzing Harness: `visi` vs. Microsoft Excel

This project uses Python scripting to generate an Excel
file, execute it in both `visi` and a real copy of Excel,
write the results to the same file, and compare the
resulting file.

---

## Setup & Requirements

### 1. Requirements

- Python 3.9+, from the project venv — **not** system Python, since
  `maturin develop` installs into whichever venv is active:
  ```bash
  source fuzz/venv/bin/activate
  pip install -r fuzz/requirements.txt
  ```
- The `visi_core` bindings (see below), and/or a compiled `visi` binary
  (`cargo build --release`)
- Microsoft Excel (macOS or Windows) for actual Excel execution.

### 2. In-process bindings

The scripts drive visi through `visi-python`,
a pyo3 extension exposing `visi-core` bindings directly
to Python.

```bash
source fuzz/venv/bin/activate
maturin develop -m visi-python/Cargo.toml --release
```

Rebuild after any change to `visi-core`.

You can configure `--backend {auto,bindings,subprocess}` for each
fuzzer.
The default `auto` uses the bindings when `import visi_core` succeeds and falls back to the CLI
with a warning otherwise.
`subprocess` runs `visi` as a CLI rather than using
the Python bindings.

---

## Usage

### Run Differential Fuzzing

```bash
python fuzz/fuzz_excel.py --excel-path "/Applications/Microsoft Excel.app" --iterations 20

python fuzz/fuzz_excel.py --driver win32com --iterations 50

python fuzz/fuzz_excel.py --driver mock --iterations 5
```

## Reproducing & Debugging Failures

When a test iteration fails, the harness automatically creates
a reproduction folder under
`fuzz_results/failures/fail_iter_<N>_seed_<SEED>/`:

```
fuzz_results/failures/fail_iter_3_seed_48291/
├── source.xlsx      # Original generated workbook before evaluation
├── visi_out.xlsx    # Workbook evaluated and saved by visi
└── excel_out.xlsx   # Workbook evaluated and saved by Microsoft Excel
```

You can inspect `source.xlsx` or run `visi eval source.xlsx --output test.xlsx` to debug formula evaluation discrepancies directly.
