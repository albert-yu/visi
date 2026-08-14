#!/usr/bin/env python3
"""visi execution drivers, shared by fuzz_excel / fuzz_pivot / fuzz_chart.

Each driver has two interchangeable backends:

  "bindings"    -- the `visi_core` extension module (the `visi-python` crate,
                   built with `maturin develop`). One process, and no disk
                   round trip per operation.
  "subprocess"  -- the `visi` CLI, exactly as before.

The subprocess backend is kept for two reasons, and the second is the
important one:

  1. It still works on a checkout where the extension module hasn't been
     built.
  2. It is the crash-triage mode. Under "bindings" the engine shares this
     process, so a Rust panic is a catchable PanicException but an abort or a
     stack overflow (plausible -- the formula parser is recursive descent and
     the generator emits deeply nested expressions) takes the whole run down,
     losing every iteration's progress. Under "subprocess" it costs one
     iteration.

The two backends must stay observationally identical. That is not an
aspiration, it is checked: see fuzz/test_backend_parity.py.
"""

import json
import os
import shutil
import subprocess

try:
    import visi_core as _vc

    _IMPORT_ERROR = None
except ImportError as exc:  # not built; the CLI backend still works
    _vc = None
    _IMPORT_ERROR = exc

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# The `visi` CLI has no timeout today, so a hang stalls a whole run forever.
# Generous enough that a slow-but-progressing workbook is never cut off.
CLI_TIMEOUT_SECONDS = 120


def bindings_available():
    """Whether `import visi_core` succeeded."""
    return _vc is not None


def bindings_hint():
    """Why the bindings aren't in use, and how to fix it."""
    return (
        f"visi_core bindings unavailable ({_IMPORT_ERROR}); falling back to the visi CLI.\n"
        f"  Build them with:  source fuzz/venv/bin/activate && "
        f"pip install -r fuzz/requirements.txt && "
        f"maturin develop -m visi-python/Cargo.toml --release"
    )


def resolve_visi_binary(binary_path=None):
    """The explicit `--visi-path` if it exists, else the newer of
    target/release/visi and target/debug/visi.

    Preferring the newer of the two is deliberate: it means a `cargo build`
    (debug) during development is picked up without having to remember to pass
    a path, at the cost of a much slower binary.
    """
    if binary_path and os.path.exists(binary_path):
        return binary_path
    candidates = [
        os.path.join(PROJECT_ROOT, "target", flavor, "visi")
        for flavor in ("release", "debug")
    ]
    existing = [p for p in candidates if os.path.exists(p)]
    if not existing:
        return binary_path
    return max(existing, key=os.path.getmtime)


def pick_backend(requested):
    """Resolves 'auto' | 'bindings' | 'subprocess' to a concrete backend."""
    if requested == "bindings":
        if not bindings_available():
            raise RuntimeError(bindings_hint())
        return "bindings"
    if requested == "subprocess":
        return "subprocess"
    return "bindings" if bindings_available() else "subprocess"


def add_backend_arg(parser):
    """Adds the shared --backend flag to a fuzzer's ArgumentParser."""
    parser.add_argument(
        "--backend",
        choices=["auto", "bindings", "subprocess"],
        default="auto",
        help=(
            "How to drive visi: in-process bindings, the CLI, or (default) "
            "whichever is available. Use 'subprocess' to triage a crash."
        ),
    )


class _BaseDriver:
    def __init__(self, binary_path=None, backend="auto"):
        self.binary_path = resolve_visi_binary(binary_path)
        self.backend = pick_backend(backend)

    def describe(self):
        """What to print in a run banner."""
        if self.backend == "bindings":
            return f"bindings ({_vc.__file__})"
        return f"subprocess ({self.binary_path})"

    def _cli(self, subcommand, args):
        cmd = [self.binary_path, subcommand] + args
        res = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=CLI_TIMEOUT_SECONDS,
        )
        if res.returncode != 0:
            raise RuntimeError(
                f"visi {subcommand} {' '.join(args)} failed with code "
                f"{res.returncode}:\nSTDOUT: {res.stdout}\nSTDERR: {res.stderr}"
            )
        return res.stdout


class VisiDriver(_BaseDriver):
    """Recalculates a workbook's formulas and writes the result.

    The constructor signature is preserved for reverse_engineer_financial.py,
    which builds it as VisiDriver(binary_path) and calls run(src, dst).
    """

    def run(self, input_file, output_file):
        """Evaluates `input_file` into `output_file`.

        Returns the .xlsx bytes written, so a caller can parse them without a
        second read. The file is written either way: a failure artifact is
        worth far more than the millisecond saved by skipping it.
        """
        if self.backend == "bindings":
            wb = _vc.Workbook.load(input_file)
            wb.evaluate()
            data = wb.save_bytes()
            with open(output_file, "wb") as f:
                f.write(data)
            return data

        self._cli("eval", [input_file, "--output", output_file])
        with open(output_file, "rb") as f:
            return f.read()


class VisiChartDriver(_BaseDriver):
    """Adds a chart, then edits it, mirroring `visi chart add` + `chart edit`."""

    def run(self, source_file, range_str, add_config, edit_config, output_file):
        if self.backend == "subprocess":
            return self._run_cli(
                source_file, range_str, add_config, edit_config, output_file
            )

        wb = _vc.Workbook.load(source_file)
        wb.add_chart(
            "Sheet1", add_config["chart_type"], range_str, title=add_config["title"] or None
        )

        # `chart add -i` wrote the file and `chart list --json` reopened it.
        # Keep that round trip: it is the only differential coverage of the
        # chart OOXML writer/reader pair. The id has to be re-read afterwards
        # because import re-derives it from the sheet name and the chart's
        # position, so add_chart's return value is stale by now.
        wb = wb.roundtrip()
        chart_id = wb.charts()[0]["id"]

        wb.edit_chart(
            chart_id,
            chart_type=edit_config["chart_type"],
            title=edit_config["title"] or None,
            clear_title=not edit_config["title"],
            xlabel=edit_config["xlabel"] or None,
            clear_xlabel=not edit_config["xlabel"],
            ylabel=edit_config["ylabel"] or None,
            clear_ylabel=not edit_config["ylabel"],
            show_legend=bool(edit_config["show_legend"]),
        )
        wb.save(output_file)

    def _run_cli(self, source_file, range_str, add_config, edit_config, output_file):
        shutil.copyfile(source_file, output_file)

        add_args = [
            "add", output_file,
            "--sheet", "Sheet1",
            "--chart-type", add_config["chart_type"],
            "--range", range_str,
            "-i",
        ]
        if add_config["title"]:
            add_args += ["--title", add_config["title"]]
        self._cli("chart", add_args)

        # `chart add` has no --json output of its own; look the new chart's id
        # up via `chart list --json` (the only chart in the file).
        charts = json.loads(self._cli("chart", ["list", output_file, "--json"]))
        chart_id = charts[0]["id"]

        edit_args = [
            "edit", output_file,
            "--id", str(chart_id),
            "--chart-type", edit_config["chart_type"],
            "-i",
        ]
        if edit_config["title"]:
            edit_args += ["--title", edit_config["title"]]
        else:
            edit_args.append("--clear-title")
        if edit_config["xlabel"]:
            edit_args += ["--xlabel", edit_config["xlabel"]]
        else:
            edit_args.append("--clear-xlabel")
        if edit_config["ylabel"]:
            edit_args += ["--ylabel", edit_config["ylabel"]]
        else:
            edit_args.append("--clear-ylabel")
        edit_args.append("--show-legend" if edit_config["show_legend"] else "--hide-legend")
        self._cli("chart", edit_args)


class VisiPivotDriver(_BaseDriver):
    """Builds a pivot table field by field, mirroring the `visi pivot` verbs."""

    def run(self, source_file, config, output_file, pivot_name, dest_cell, dest_rc):
        """`dest_cell` is A1 for the CLI path; `dest_rc` is the 0-based (row,
        col) equivalent for the bindings path. Both come from the caller so
        that neither this module nor the bindings needs an A1 parser."""
        if self.backend == "subprocess":
            return self._run_cli(source_file, config, output_file, pivot_name, dest_cell)

        dest_row, dest_col = dest_rc
        wb = _vc.Workbook.load(source_file)

        if config["table_name"]:
            wb.add_pivot_from_table(
                pivot_name,
                config["table_name"],
                dest_row=dest_row,
                dest_col=dest_col,
                grand_totals_row=config["grand_totals_row"],
                grand_totals_col=config["grand_totals_col"],
            )
        else:
            sr, sc, er, ec = config["source_bounds"]
            wb.add_pivot_from_range(
                pivot_name,
                start_row=sr, start_col=sc, end_row=er, end_col=ec,
                dest_row=dest_row, dest_col=dest_col,
                grand_totals_row=config["grand_totals_row"],
                grand_totals_col=config["grand_totals_col"],
            )
        # One roundtrip per mutation, standing in for the file each `-i` CLI
        # invocation used to write and reopen. Dropping these would quietly
        # stop exercising pivot_xlsx.rs's hand-rolled OOXML.
        wb = wb.roundtrip()

        for area, fields in (("row", config["row_fields"]), ("column", config["col_fields"])):
            for f in fields:
                wb.add_pivot_field(pivot_name, area, f["column"], subtotal=f["subtotal"])
                wb = wb.roundtrip()

        for f in config["value_fields"]:
            wb.add_pivot_field(pivot_name, "value", f["column"], agg=f["agg"])
            wb = wb.roundtrip()

        if config["filter_field"]:
            column = config["filter_field"]["column"]
            wb.add_pivot_field(pivot_name, "filter", column)
            wb = wb.roundtrip()
            # Deliberately the LAST mutation, with no roundtrip after it:
            # PivotFilterField.selected_values is not reconstructed on import,
            # so a round trip here would silently reset the filter to "all".
            #
            # The empty-list guard is defensive: PivotFuzzGenerator no longer
            # emits an empty selection (it means "select nothing", which real
            # Excel cannot represent -- see the comment beside `selected` in
            # fuzz_pivot.py), but a hand-written config still can, and
            # applying one would compare visi's empty grid against Excel's
            # full one. Leave such a field unfiltered, as BuildFuzzPivot.bas
            # and the CLI backend below both do. The engine's empty-selection
            # behavior is covered directly, in
            # test_empty_filter_selection_is_bindings_only.
            values = config["filter_field"]["values"]
            if values:
                wb.set_pivot_filter(pivot_name, column, values)

        wb.save(output_file)

    def _run_cli(self, source_file, config, output_file, pivot_name, dest_cell):
        shutil.copyfile(source_file, output_file)

        create_args = ["create", output_file, "--name", pivot_name, "--dest", dest_cell, "-i"]
        if config["table_name"]:
            create_args += ["--source-table", config["table_name"]]
        else:
            create_args += ["--source-range", config["source_range"]]
        if not config["grand_totals_row"]:
            create_args.append("--no-grand-totals-row")
        if not config["grand_totals_col"]:
            create_args.append("--no-grand-totals-col")
        self._cli("pivot", create_args)

        for area, fields in (("row", config["row_fields"]), ("column", config["col_fields"])):
            for f in fields:
                args = [
                    "add-field", output_file, "--name", pivot_name,
                    "--area", area, "--column", f["column"], "-i",
                ]
                if not f["subtotal"]:
                    args.append("--no-subtotal")
                self._cli("pivot", args)

        for f in config["value_fields"]:
            self._cli(
                "pivot",
                ["add-field", output_file, "--name", pivot_name, "--area", "value",
                 "--column", f["column"], "--agg", f["agg"], "-i"],
            )

        if config["filter_field"]:
            self._cli(
                "pivot",
                ["add-field", output_file, "--name", pivot_name, "--area", "filter",
                 "--column", config["filter_field"]["column"], "-i"],
            )
            values = config["filter_field"]["values"]
            if values:
                self._cli(
                    "pivot",
                    ["filter", output_file, "--name", pivot_name,
                     "--column", config["filter_field"]["column"],
                     "--values", ",".join(values), "-i"],
                )
            # else: the config wants "select nothing", which this backend
            # cannot express -- the CLI's `filter` verb takes a comma list or
            # --clear, with no verb for an empty selection. Leave the field
            # unfiltered, same as the bindings path and BuildFuzzPivot.bas.
