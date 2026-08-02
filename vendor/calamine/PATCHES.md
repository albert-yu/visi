# Vendored patches to calamine 0.26.1

This is a trimmed, patched copy of [calamine](https://github.com/tafia/calamine)
0.26.1 (library source only -- examples/benches/tests dropped, see this
directory's `Cargo.toml`). Wired in via `[patch.crates-io]` in the workspace
root `Cargo.toml`, so `libvisi`'s `calamine = "0.26"` dependency resolves to
this copy transparently.

## Why vendor instead of just upgrading

Confirmed the bugs below are still present unpatched in calamine's latest
release (0.36.1 at the time of writing) -- upgrading the version pin
wouldn't fix them, and jumping ten minor versions carries its own
(unrelated) API-compatibility risk. Vendoring small, targeted patches is
the smaller change.

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

## Patch: worksheet-to-table relationships with absolute package-path targets resolve to nothing

**File:** `src/xlsx/mod.rs`, `Xlsx::read_table_metadata`.

**Symptom:** an Excel Table written by `openpyxl` (or any other writer that
emits absolute-path relationship targets) silently imports as zero tables
-- `workbook.load_tables()` succeeds but finds none, with no error. Found
via `libvisi`'s xlsx import path on an untouched `openpyxl`-authored
`.xlsx` file (`visi table list` reported none); see GitHub issue #16.

**Root cause:** `read_table_metadata` reads each worksheet's
`_rels/sheetN.xml.rels` and resolves the `Target` of each `Relationship`
whose `Type` is the table relationship type. It special-cases `../`-relative
targets (the form real Excel and `rust_xlsxwriter` always emit) by
resolving them against the worksheet's parent folder, and skips empty
targets, but otherwise pushes the `Target` string through unchanged. OPC
also permits absolute package-path targets like `/xl/tables/table1.xml`
(what `openpyxl` writes), and those fall into that unchanged-passthrough
branch. Zip entry names never have a leading `/`, so `xml_reader` can't
find a matching entry for `/xl/tables/table1.xml`, returns `None`, and the
table location is silently dropped (`continue`) instead of erroring.

**Fix:** in that branch, strip a leading `/` before pushing the path, so
`/xl/tables/table1.xml` resolves the same as `xl/tables/table1.xml` would.

## Upstream

Worth filing as a PR against `tafia/calamine` -- both patches above are
minimal, clearly scoped fixes with obvious regression tests. If merged,
this vendor directory can be dropped once `libvisi` upgrades past whatever
release includes them. Not filed yet.
