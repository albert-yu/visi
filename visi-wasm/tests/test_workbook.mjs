import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const pkgDir = path.resolve(__dirname, "..", "pkg");
const wasmPath = path.join(pkgDir, "visi_wasm_bg.wasm");
const jsPath = path.join(pkgDir, "visi_wasm.js");

const wasmBytes = fs.readFileSync(wasmPath);
const visi = await import(jsPath);
visi.initSync({ module: wasmBytes });

const {
  Workbook,
  CellError,
  evalBytes,
  detectWorkbookFormat,
  checkSyntax,
  checkSyntaxPartial,
  runMacro,
  colIdxToLetters,
  colLettersToIdx,
  parseCellRef,
  parseRangeRef,
  isCellError,
  version,
} = visi;

test("version is non-empty string", () => {
  assert.equal(typeof version(), "string");
  assert.ok(version().length > 0);
});

test("new empty workbook has one sheet", () => {
  const wb = new Workbook();
  assert.deepEqual(wb.sheetNames, ["Sheet1"]);
  assert.deepEqual(wb.dimensions(), { rows: 10, cols: 5 });
  wb.free();
});

test("evaluate and read back", () => {
  const wb = new Workbook();
  wb.setCell(0, 0, "10");
  wb.setCell(0, 1, "=A1*2");
  wb.evaluate();
  assert.equal(wb.getCell(0, 0), 10);
  assert.equal(wb.getCell(0, 1), 20);
  wb.free();
});

test("toBytes and fromBytes roundtrip", () => {
  const wb = new Workbook();
  wb.setCell(0, 0, "42");
  wb.evaluate();

  const bytes = wb.toBytes();
  assert.ok(bytes instanceof Uint8Array);
  assert.equal(detectWorkbookFormat(bytes), "xlsx");

  const again = Workbook.fromBytes(bytes);
  assert.deepEqual(again.sheetNames, wb.sheetNames);
  assert.equal(again.getCell(0, 0), 42);
  wb.free();
  again.free();
});

test("xlsb saveBytes and loadBytes roundtrip", () => {
  const wb = new Workbook();
  wb.setCell(0, 0, "7");
  wb.setCell(0, 1, "=A1+5");
  wb.evaluate();

  const bytes = wb.saveBytes("xlsb");
  assert.equal(detectWorkbookFormat(bytes), "xlsb");

  const again = Workbook.loadBytes(bytes);
  assert.equal(again.getCell(0, 0), 7);
  assert.equal(again.getCell(0, 1), 12);
  wb.free();
  again.free();
});

test("roundtrip matches manual save and load", () => {
  const wb = new Workbook();
  wb.setCell(0, 0, "7");
  wb.evaluate();
  const rt = wb.roundtrip();
  const manual = Workbook.fromBytes(wb.toBytes());
  assert.equal(rt.getCell(0, 0), manual.getCell(0, 0));
  wb.free();
  rt.free();
  manual.free();
});

test("evalBytes evaluates in one shot", () => {
  const wb = new Workbook();
  wb.setCell(0, 0, "3");
  wb.setCell(0, 1, "=A1+4");
  const rawBytes = wb.toBytes();
  wb.free();

  const evaluatedBytes = evalBytes(rawBytes);
  const loaded = Workbook.fromBytes(evaluatedBytes);
  assert.equal(loaded.getCell(0, 1), 7);
  loaded.free();
});

test("result types map to native JS primitives", () => {
  const cases = [
    ["10", 10, "number"],
    ["=1+1", 2.0, "number"],
    ["=1.5+1", 2.5, "number"],
    ['=CONCATENATE("a","b")', "ab", "string"],
    ["=TRUE()", true, "boolean"],
  ];
  for (const [src, expected, expectedType] of cases) {
    const wb = new Workbook();
    wb.setCell(0, 0, src);
    wb.evaluate();
    const v = wb.getCell(0, 0);
    assert.equal(v, expected);
    assert.equal(typeof v, expectedType);
    wb.free();
  }
});

test("typecheck returns structured diagnostics and supports coords or A1 selectors", () => {
  const wb = new Workbook();
  wb.setCell(0, 0, '="abc"+1');
  wb.setCell(0, 1, '="1"+1');
  wb.setCell(0, 2, "=1+2");

  const result = wb.typecheck();
  assert.equal(result.error_count, 1);
  assert.equal(result.warning_count, 1);
  assert.equal(result.errorCount, 1);
  assert.equal(result.warningCount, 1);
  assert.equal(result.diagnostics.length, 2);
  assert.deepEqual(result.diagnostics[0], {
    sheet: "Sheet1",
    cell: "A1",
    row: 0,
    col: 0,
    severity: "error",
    kind: "type",
    message: "left operand expects a number, got text",
    formula: '="abc"+1',
  });

  const wb2 = new Workbook();
  wb2.setCell(0, 0, "=NO_SUCH_FN(1)");
  wb2.setCell(1, 1, "=Missing!A1");

  const cellResult = wb2.typecheck({ cell: [1, 1] });
  assert.equal(cellResult.errorCount, 1);
  assert.equal(cellResult.diagnostics[0].cell, "B2");
  assert.equal(cellResult.diagnostics[0].kind, "reference");

  const a1CellResult = wb2.typecheck({ cell: "B2" });
  assert.equal(a1CellResult.errorCount, 1);
  assert.equal(a1CellResult.diagnostics[0].cell, "B2");

  const rangeResult = wb2.typecheck({ range: [0, 0, 0, 1] });
  assert.equal(rangeResult.errorCount, 1);
  assert.equal(rangeResult.diagnostics[0].cell, "A1");
  assert.equal(rangeResult.diagnostics[0].kind, "function");

  const a1RangeResult = wb2.typecheck({ range: "A1:B1" });
  assert.equal(a1RangeResult.errorCount, 1);
  assert.equal(a1RangeResult.diagnostics[0].cell, "A1");

  assert.throws(
    () => wb2.typecheck({ cell: [0, 0], range: [0, 0, 1, 1] }),
    (err) => err.name === "InvalidArgumentError"
  );
  assert.throws(
    () => wb2.typecheck({ range: [1, 0, 0, 0] }),
    (err) => err.name === "InvalidArgumentError"
  );

  wb.free();
  wb2.free();
});

test("setCell with explicit cellType and setCellType on existing cell", () => {
  const wb = new Workbook();
  wb.setCell(0, 0, "12345", undefined, "string");
  wb.setCell(0, 1, "12345");
  wb.evaluate();
  assert.equal(wb.getCellType(0, 0), "string");
  assert.equal(wb.getCell(0, 0), "12345");
  assert.equal(wb.getCellType(0, 1), "int");
  assert.equal(wb.getCell(0, 1), 12345);

  const again = wb.roundtrip();
  assert.equal(again.getCellType(0, 0), "string");
  assert.equal(again.getCell(0, 0), "12345");

  const wb2 = new Workbook();
  wb2.setCell(0, 0, "007");
  wb2.setCellType(0, 0, "string");
  wb2.evaluate();
  assert.equal(wb2.getCellType(0, 0), "string");
  assert.equal(wb2.getCell(0, 0), "007");

  wb.free();
  again.free();
  wb2.free();
});

test("blank cell is null and error value is CellError", () => {
  const wb = new Workbook();
  wb.setCell(0, 0, "=1/0");
  wb.setCell(0, 1, '=CONCATENATE("#DIV/0!")');
  wb.evaluate();

  assert.equal(wb.getCell(5, 5), null);

  const err = wb.getCell(0, 0);
  const text = wb.getCell(0, 1);
  assert.ok(err instanceof CellError);
  assert.equal(isCellError(err), true);
  assert.equal(isCellError(text), false);
  assert.equal(err.code, "#DIV/0!");
  assert.equal(err.toString(), "#DIV/0!");
  assert.deepEqual(err.toJSON(), { type: "error", code: "#DIV/0!" });
  assert.equal(typeof text, "string");
  wb.free();
});

test("getDisplay and getSrc", () => {
  const wb = new Workbook();
  wb.setCell(0, 0, "6/22/26");
  wb.setCell(0, 1, "=1+1");
  wb.evaluate();

  assert.equal(typeof wb.getCell(0, 0), "number");
  assert.ok(wb.getDisplay(0, 0).includes("26"));
  assert.equal(wb.getSrc(0, 1), "=1+1");
  assert.equal(wb.getCell(0, 1), 2);
  wb.free();
});

test("structured errors carry name, code, and variant metadata", () => {
  const wb = new Workbook();
  assert.throws(
    () => wb.sheetIndex("nope"),
    (err) => {
      assert.equal(err.name, "NotFoundError");
      assert.equal(err.code, "NotFoundError");
      assert.equal(err.kind, "sheet");
      assert.equal(err.objectName, "nope");
      assert.deepEqual(err.available, ["Sheet1"]);
      assert.equal(err.message, "sheet 'nope' not found. Available sheets: Sheet1");
      return true;
    }
  );

  assert.throws(
    () => wb.addChart("Sheet1", "doughnut", "Sheet1!A1:B2"),
    (err) => {
      assert.equal(err.name, "InvalidArgumentError");
      assert.ok(err.message.includes("doughnut"));
      return true;
    }
  );
  wb.free();
});

test("sheet CRUD, row/col structural edits, and circular reference detection", () => {
  const wb = new Workbook();
  wb.addSheet("Data");
  assert.deepEqual(wb.sheetNames, ["Sheet1", "Data"]);
  wb.renameSheet("Data", "Metrics");
  assert.deepEqual(wb.sheetNames, ["Sheet1", "Metrics"]);

  wb.setCell(0, 0, "10");
  wb.setCell(1, 0, "=A1*3");
  wb.evaluate();
  assert.equal(wb.getCell(1, 0), 30);

  wb.insertRow(0);
  assert.equal(wb.getSrc(2, 0), "=A2*3");
  assert.equal(wb.getCell(2, 0), 30);

  wb.deleteRow(0);
  assert.equal(wb.getSrc(1, 0), "=A1*3");

  wb.insertCol(0);
  assert.equal(wb.getSrc(1, 1), "=B1*3");
  wb.deleteCol(0);
  assert.equal(wb.getSrc(1, 0), "=A1*3");

  wb.deleteSheet("Metrics");
  assert.deepEqual(wb.sheetNames, ["Sheet1"]);

  assert.throws(
    () => wb.deleteSheet("Sheet1"),
    (err) => err.name === "LastSheetError"
  );

  wb.setCell(0, 0, "=B1+1");
  wb.setCell(0, 1, "=A1+1");
  assert.deepEqual(wb.detectCircularReferences(), [["Sheet1!A1", "Sheet1!B1"]]);
  wb.free();
});

test("CSV import and batch range / style operations", () => {
  const csvBytes = new TextEncoder().encode("item,price\nApple,1.5\nBanana,2.5\n");
  const wb = Workbook.fromCsvBytes(csvBytes, "Groceries");
  wb.evaluate();
  assert.deepEqual(wb.sheetNames, ["Groceries"]);
  assert.equal(wb.getCell(1, 1, "Groceries"), 1.5);

  wb.setCells(
    [
      {
        row: 3,
        col: 0,
        value: "Total",
        sheet: "Groceries",
        style: { bold: true, fontColor: "#FF0000" },
      },
      {
        row: 3,
        col: 1,
        value: "=SUM(B2:B3)",
        sheet: "Groceries",
        style: { numFormat: "0.00" },
      },
    ],
    true
  );

  assert.equal(wb.getCell(3, 1, "Groceries"), 4);
  assert.deepEqual(wb.getCellStyle(3, 0, "Groceries"), {
    bold: true,
    fontColor: "#FF0000",
  });

  const grid = wb.getRange(0, 0, 3, 1, "Groceries");
  assert.equal(grid.length, 4);
  assert.equal(grid[3][0].value, "Total");
  assert.equal(grid[3][0].style.bold, true);
  assert.equal(grid[3][1].value, 4);
  assert.equal(grid[3][1].src, "=SUM(B2:B3)");
  wb.free();
});

test("Excel tables CRUD and structured references", () => {
  const wb = new Workbook();
  wb.setCells(
    [
      { row: 0, col: 0, value: "Product" },
      { row: 0, col: 1, value: "Qty" },
      { row: 1, col: 0, value: "Pen" },
      { row: 1, col: 1, value: "10" },
      { row: 2, col: 0, value: "Book" },
      { row: 2, col: 1, value: "5" },
    ],
    true
  );

  const tid = wb.addTable({
    name: "Inventory",
    startRow: 0,
    startCol: 0,
    endRow: 2,
    endCol: 1,
    styleName: "TableStyleMedium2",
  });
  assert.equal(typeof tid, "number");
  assert.equal(wb.tables().length, 1);
  assert.equal(wb.tables()[0].name, "Inventory");
  assert.equal(wb.getTableStyle("Inventory"), "TableStyleMedium2");

  wb.setCell(0, 3, "=SUM(Inventory[Qty])");
  wb.evaluate();
  assert.equal(wb.getCell(0, 3), 15);

  wb.renameTableColumn("Inventory", 1, "Quantity");
  assert.equal(wb.getSrc(0, 3), "=SUM(Inventory[Quantity])");
  assert.equal(wb.getCell(0, 3), 15);

  wb.renameTable("Inventory", "Stock");
  assert.equal(wb.getSrc(0, 3), "=SUM(Stock[Quantity])");

  wb.deleteTable("Stock");
  assert.equal(wb.tables().length, 0);
  wb.free();
});

test("charts add, edit, roundtrip, and delete", () => {
  const wb = new Workbook();
  const sheet = wb.sheetNames[0];
  wb.setCells(
    [
      { row: 0, col: 0, value: "a" },
      { row: 0, col: 1, value: "1" },
      { row: 1, col: 0, value: "b" },
      { row: 1, col: 1, value: "2" },
    ],
    true
  );

  const cid = wb.addChart(sheet, "column", `${sheet}!A1:B2`, { title: "Before" });
  assert.equal(wb.charts().length, 1);
  assert.equal(wb.charts()[0].id, cid);
  assert.equal(wb.charts()[0].title, "Before");

  wb.editChart(cid, { title: "After", ylabel: "Amount", chartType: "line" });
  assert.equal(wb.charts()[0].title, "After");
  assert.equal(wb.charts()[0].ylabel, "Amount");
  assert.equal(wb.charts()[0].type, "Line");

  wb.editChart(cid, { clearTitle: true });
  assert.equal(wb.charts()[0].title, null);

  assert.throws(
    () => wb.editChart(cid, { title: "x", clearTitle: true }),
    (err) => err.name === "InvalidArgumentError"
  );

  wb.deleteChart(cid);
  assert.equal(wb.charts().length, 0);
  wb.free();
});

test("pivot tables from range, subtotals, and filters", () => {
  const wb = new Workbook();
  const rows = [
    ["Region", "Product", "Amount"],
    ["East", "Widget", "10"],
    ["East", "Gadget", "20"],
    ["West", "Widget", "30"],
    ["West", "Gadget", "40"],
  ];
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      wb.setCell(r, c, rows[r][c]);
    }
  }
  wb.evaluate();

  wb.addPivotFromRange("P", {
    startRow: 0,
    startCol: 0,
    endRow: rows.length - 1,
    endCol: 2,
    destRow: 0,
    destCol: 5,
  });
  wb.addPivotField("P", "row", "Region", { subtotal: false });
  wb.addPivotField("P", "row", "Product");
  wb.addPivotField("P", "value", "Amount", { agg: "sum" });

  const p = wb.pivots()[0];
  assert.equal(p.name, "P");
  assert.deepEqual(p.rowFields, ["Region", "Product"]);
  assert.deepEqual(p.valueFields, ["Sum of Amount"]);
  assert.equal(p.subtotals.Region, false);
  assert.equal(p.subtotals.Product, true);

  wb.addPivotField("P", "filter", "Region");
  wb.setPivotFilter("P", "Region", []);
  assert.deepEqual(wb.pivots()[0].filterSelections.Region, []);

  wb.setPivotFilter("P", "Region", ["East"]);
  assert.deepEqual(wb.pivots()[0].filterSelections.Region, ["East"]);

  wb.setPivotFilter("P", "Region", null);
  assert.equal(wb.pivots()[0].filterSelections.Region, null);
  wb.free();
});

test("VBA macros checkSyntax, add, run, roundtrip, and remove", () => {
  const macroSrc =
    'Attribute VB_Name = "Runner"\n' +
    "Public Function Go() As Variant\n" +
    '    Range("A1").Value = 4\n' +
    '    Range("B1").Formula = "=A1*2"\n' +
    '    Range("C1").Value = Application.WorksheetFunction.Sum(Range("A1:B1"))\n' +
    '    Go = Range("C1").Value\n' +
    "End Function\n";

  assert.deepEqual(checkSyntax(macroSrc), ["Go"]);
  assert.deepEqual(checkSyntaxPartial("Sub Caller()\n    ExternalProc 1\nEnd Sub\n"), ["Caller"]);

  const standalone = runMacro(
    "Function Add(a, b)\n    Add = a + b\nEnd Function\n",
    "Add",
    ["3", "4"]
  );
  assert.equal(standalone.value, "7");

  const wb = new Workbook();
  assert.equal(wb.hasMacros(), false);
  wb.addMacro("Runner", macroSrc);
  assert.equal(wb.hasMacros(), true);
  assert.equal(wb.macros().length, 1);

  const out = wb.runMacro("Go");
  assert.equal(out.value, "12");
  assert.equal(out.mutated, true);
  assert.equal(wb.getCell(0, 2), 12);

  const rt = wb.roundtrip();
  assert.equal(rt.macros().length, 1);
  assert.equal(rt.macros()[0].name, "Runner");

  wb.removeMacro("Runner");
  assert.equal(wb.macros().length, 0);
  wb.free();
  rt.free();
});

test("locale and A1 coordinate helpers", () => {
  const wbUs = new Workbook("en-US");
  assert.equal(wbUs.locale, "en-US");
  wbUs.setCell(0, 0, "06/07/2026");
  wbUs.evaluate();
  assert.equal(wbUs.getCell(0, 0), 46180);

  const wbGb = new Workbook("en-GB");
  wbGb.setCell(0, 0, "06/07/2026");
  wbGb.evaluate();
  assert.equal(wbGb.getCell(0, 0), 46209);

  assert.equal(colIdxToLetters(0), "A");
  assert.equal(colIdxToLetters(26), "AA");
  assert.equal(colLettersToIdx("AA"), 26);
  assert.deepEqual(parseCellRef("Sheet1!B3"), { sheet: "Sheet1", row: 2, col: 1 });
  assert.deepEqual(parseRangeRef("A1:C5"), {
    sheet: null,
    startRow: 0,
    startCol: 0,
    endRow: 4,
    endCol: 2,
  });

  wbUs.free();
  wbGb.free();
});
