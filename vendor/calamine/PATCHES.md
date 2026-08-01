# Vendored patches to calamine 0.26.1

This is a trimmed, patched copy of [calamine](https://github.com/tafia/calamine)
0.26.1 (library source only -- examples/benches/tests dropped, see this
directory's `Cargo.toml`). Wired in via `[patch.crates-io]` in the workspace
root `Cargo.toml`, so `libvisi`'s `calamine = "0.26"` dependency resolves to
this copy transparently.

## Why vendor instead of just upgrading

Confirmed the bug below is still present unpatched in calamine's latest
release (0.36.1 at the time of writing) -- upgrading the version pin
wouldn't fix it, and jumping ten minor versions carries its own (unrelated)
API-compatibility risk. Vendoring a one-function patch is the smaller,
more targeted change.

## Patch: `Range::range` panics on a table with zero data rows

**File:** `src/lib.rs`, `Range::range`.

**Symptom:** exporting an Excel Table with a header row but zero data rows,
then reimporting it, panics inside `calamine::Xlsx::table_by_name` with
"invalid range bounds" (found via `libvisi`'s pivot-table fuzz testing --
see `fuzz/README.md`'s "Known caveats" section).

**Root cause:** `Range::range(start, end)` unconditionally calls
`Range::new(start, end)` first, which panics if `start > end`. Reachable
via `Xlsx::table_by_name`/`table_by_name_ref`, which call
`range.range(start, end)` with a table's data-only `Dimensions`.
`read_table_metadata`'s header-row adjustment (`dims.start.0 +=
header_row_count`) pushes `start.0` past `end.0` with no clamping when a
table has a header row but zero data rows -- producing exactly the invalid
`start > end` pair `range()` mishandles.

**Fix:** guard `Range::range` against `start.0 > end.0 || start.1 > end.1`
and return `Range::empty()` instead of constructing an invalid `Range`.
This is the lowest common layer -- fixes both `table_by_name` and
`table_by_name_ref` (and any other caller that hits the same degenerate
bounds) without touching call sites, and matches `range()`'s own
"no overlap" fallback a few lines below (which already returns gracefully
rather than panicking; the panic happens before that check is ever
reached).

## Upstream

Worth filing as a PR against `tafia/calamine` -- it's a minimal, clearly
scoped fix with an obvious regression test. If merged, this vendor
directory can be dropped once `libvisi` upgrades past whatever release
includes it. Not filed yet.
