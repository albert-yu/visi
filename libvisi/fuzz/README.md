# VBA parsing fuzz targets

`cargo-fuzz` (libFuzzer) harness for the two places in `libvisi` that parse
completely untrusted bytes -- a `.xlsm`/`.bin` someone else authored, not
anything this codebase produced itself. This is a different kind of fuzzing
than `../../fuzz/` at the repo root: that one is differential (compares
`visi`'s formula/pivot evaluation against real Excel's); this one just hunts
for panics, unbounded allocation, and infinite loops on malformed input.

## Targets

- **`ovba_decompress`** -- `core::ovba::decompress`, the MS-OVBA "Compressed
  Container" LZ77 decoder. Every module's source stream and the project's
  `dir` stream go through this.
- **`vba_import`** -- `core::vba_xlsx::parse_vba_project_from_cfb_bytes`, the
  full import pipeline: CFB container parsing, `dir`-stream decompression,
  `PROJECTMODULES` record walking, and per-module stream decompression. This
  is what runs on `xl/vbaProject.bin` the moment any `.xlsm` is opened.

## Setup

```bash
rustup toolchain install nightly   # cargo-fuzz needs nightly for sanitizer instrumentation
cargo install cargo-fuzz
```

## Running

From `libvisi/`:

```bash
cargo +nightly fuzz run ovba_decompress
mkdir -p fuzz/corpus/vba_import
cargo +nightly fuzz run vba_import fuzz/corpus/vba_import fuzz/seeds/vba_import   # seed corpus, see below
```

Add `-- -max_total_time=60` (or `-runs=N`) to bound a run instead of fuzzing
indefinitely. A crash writes a reproducer to `fuzz/artifacts/<target>/` --
rerun `cargo +nightly fuzz run <target> fuzz/artifacts/<target>/<file>` to
replay it under a debugger, and minimize with `cargo +nightly fuzz tmin`.

`fuzz/corpus/` and `fuzz/artifacts/` are gitignored scratch state, rebuilt by
libFuzzer itself as it explores -- nothing to commit there.

## Seed corpus

`vba_import` first has to get past `cfb::CompoundFile::open` recognizing its
input as a valid CFB container at all (magic bytes `D0 CF 11 E0 A1 B1 1A
E1`), which random mutation from an empty corpus essentially never manages
on its own -- an unseeded run tops out around `cov: 86`. `fuzz/seeds/`
holds two real CFB-wrapped `vbaProject.bin`s (an empty synthetic project and
one with a standard module) checked into git for exactly this reason, and
passing them alongside the (gitignored) corpus dir as in the `vba_import`
command above gets libFuzzer past that gate immediately -- coverage goes
from `cov: 86` to `cov: 1600+` in the same time budget.

**Always list the gitignored `fuzz/corpus/<target>` dir first and the
checked-in `fuzz/seeds/<target>` dir after.** libFuzzer writes every new
coverage-increasing input to the *first* corpus directory it's given and
only reads the rest -- pass `fuzz/seeds/vba_import` alone (with no corpus
dir preceding it) and it becomes the write target instead, silently
flooding the checked-in seeds with hundreds of generated files on the very
first run. Regenerate the seeds themselves (distinct from the corpus) with:

```bash
cargo run -p libvisi --example dump_vba_fuzz_seeds
```

## What's covered elsewhere

`decompress`'s roundtrip property (`decompress(compress(x)) == x`) and its
never-panics guarantee on arbitrary bytes are also checked via `proptest` in
`core::ovba::tests` as part of the normal `cargo test -p libvisi` run --
useful for CI and bisecting, where a libFuzzer corpus isn't practical to
carry around. This `fuzz/` crate is for deeper, longer-running exploration
a property test's fixed small sample sizes won't reach.
