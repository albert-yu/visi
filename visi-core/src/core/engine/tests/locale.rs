use crate::core::WorkbookManager;
use crate::core::engine::{CellRef, Sheet, SheetInit};
use crate::core::locale::Locale;

#[test]
fn test_sheet_commit_date_parsing_en_us_vs_en_gb() {
    // US: 06/07/2026 = June 7, 2026
    let mut sheet_us = Sheet::new(SheetInit {
        name: Some("US".to_string()),
        rows: 5,
        cols: 5,
        ..Default::default()
    });
    sheet_us.locale = Locale::en_us();
    sheet_us.set_cell_src(0, 0, "06/07/2026".to_string());
    sheet_us.commit(None).unwrap();

    let val_us = sheet_us.get_result_data(&CellRef::new(0, 0));
    // June 7, 2026 serial = 46180
    assert_eq!(val_us.to_string(), "46180");
    assert_eq!(sheet_us.get_display_string(&CellRef::new(0, 0)), "6/7/2026");

    // GB: 06/07/2026 = July 6, 2026
    let mut sheet_gb = Sheet::new(SheetInit {
        name: Some("GB".to_string()),
        rows: 5,
        cols: 5,
        ..Default::default()
    });
    sheet_gb.locale = Locale::en_gb();
    sheet_gb.set_cell_src(0, 0, "06/07/2026".to_string());
    sheet_gb.commit(None).unwrap();

    let val_gb = sheet_gb.get_result_data(&CellRef::new(0, 0));
    // July 6, 2026 serial = 46209
    assert_eq!(val_gb.to_string(), "46209");
    assert_eq!(sheet_gb.get_display_string(&CellRef::new(0, 0)), "6/7/2026");
}

#[test]
fn test_sheet_commit_german_dates() {
    let mut sheet_de = Sheet::new(SheetInit {
        name: Some("DE".to_string()),
        rows: 5,
        cols: 5,
        ..Default::default()
    });
    sheet_de.locale = Locale::de_de();
    sheet_de.set_cell_src(0, 0, "22.06.2026".to_string());
    sheet_de.set_cell_src(1, 0, "22. Juni 2026".to_string());
    sheet_de.commit(None).unwrap();

    // 22 June 2026 serial = 46195
    assert_eq!(
        sheet_de.get_result_data(&CellRef::new(0, 0)).to_string(),
        "46195"
    );
    assert_eq!(
        sheet_de.get_result_data(&CellRef::new(1, 0)).to_string(),
        "46195"
    );
}

#[test]
fn test_sheet_commit_french_dates() {
    let mut sheet_fr = Sheet::new(SheetInit {
        name: Some("FR".to_string()),
        rows: 5,
        cols: 5,
        ..Default::default()
    });
    sheet_fr.locale = Locale::fr_fr();
    sheet_fr.set_cell_src(0, 0, "14/07/2026".to_string());
    sheet_fr.set_cell_src(1, 0, "14 juillet 2026".to_string());
    sheet_fr.commit(None).unwrap();

    // 14 July 2026 serial = 46217
    assert_eq!(
        sheet_fr.get_result_data(&CellRef::new(0, 0)).to_string(),
        "46217"
    );
    assert_eq!(
        sheet_fr.get_result_data(&CellRef::new(1, 0)).to_string(),
        "46217"
    );
}

#[test]
fn test_datevalue_formula_with_locale() {
    // US evaluates "06/07/2026" as June 7 (46180)
    let mut sheet_us = Sheet::new(SheetInit::default());
    sheet_us.locale = Locale::en_us();
    sheet_us.set_cell_src(0, 0, "=DATEVALUE(\"06/07/2026\")".to_string());
    sheet_us.commit(None).unwrap();
    assert_eq!(
        sheet_us.get_result_data(&CellRef::new(0, 0)).to_string(),
        "46180"
    );

    // GB evaluates "06/07/2026" as July 6 (46209)
    let mut sheet_gb = Sheet::new(SheetInit::default());
    sheet_gb.locale = Locale::en_gb();
    sheet_gb.set_cell_src(0, 0, "=DATEVALUE(\"06/07/2026\")".to_string());
    sheet_gb.commit(None).unwrap();
    assert_eq!(
        sheet_gb.get_result_data(&CellRef::new(0, 0)).to_string(),
        "46209"
    );

    // DE evaluates dot separator
    let mut sheet_de = Sheet::new(SheetInit::default());
    sheet_de.locale = Locale::de_de();
    sheet_de.set_cell_src(0, 0, "=DATEVALUE(\"22.06.2026\")".to_string());
    sheet_de.commit(None).unwrap();
    assert_eq!(
        sheet_de.get_result_data(&CellRef::new(0, 0)).to_string(),
        "46195"
    );
}

#[test]
fn test_value_formula_with_locale() {
    let mut sheet_de = Sheet::new(SheetInit::default());
    sheet_de.locale = Locale::de_de();
    sheet_de.set_cell_src(0, 0, "=VALUE(\"22.06.2026\")".to_string());
    sheet_de.commit(None).unwrap();
    assert_eq!(
        sheet_de.get_result_data(&CellRef::new(0, 0)).to_string(),
        "46195"
    );
}

#[test]
fn test_workbook_set_locale() {
    let mut wb = WorkbookManager::new_empty().unwrap();
    wb.set_locale(Locale::de_de());
    assert_eq!(wb.locale.code, "de-DE");
    assert_eq!(wb.sheets[0].locale.code, "de-DE");

    wb.sheets[0].set_cell_src(0, 0, "22.06.2026".to_string());
    wb.sheets[0].commit(None).unwrap();
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&CellRef::new(0, 0))
            .to_string(),
        "46195"
    );
}
