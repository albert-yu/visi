use crate::Error;
use crate::core::{Locale, Sheet, SheetInit, col_idx_to_letters};

pub(crate) fn parse_csv_sheet(
    buffer: &[u8],
    name: &str,
    id: Option<u64>,
    locale: &Locale,
) -> crate::Result<Sheet> {
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(buffer);
    let mut rows = Vec::new();
    let mut col_count = 0;

    for record in reader.records() {
        let record = record.map_err(|e| Error::Csv(e.to_string()))?;
        col_count = col_count.max(record.len());
        rows.push(record.iter().map(str::to_string).collect::<Vec<_>>());
    }

    let mut sheet = Sheet::new(SheetInit {
        id,
        name: Some(name.to_string()),
        rows: rows.len(),
        cols: col_count,
    });
    sheet.locale = locale.clone();

    for (col_idx, column) in sheet.columns.iter_mut().enumerate() {
        column.name = col_idx_to_letters(col_idx);
    }

    for (row_idx, row) in rows.into_iter().enumerate() {
        for (col_idx, value) in row.into_iter().enumerate() {
            sheet.set_cell_src(row_idx, col_idx, value);
        }
    }

    sheet.uncommitted_actions.clear();
    Ok(sheet)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::{CellRef, CellType};

    #[test]
    fn parses_records_without_treating_the_first_row_as_headers() {
        let mut sheet = parse_csv_sheet(
            b"name,value,double\nWidget,12,=B2*2\n\"Hello, world\",true,\"line1\nline2\"\n",
            "Data",
            None,
            &Locale::default(),
        )
        .unwrap();

        sheet.commit(None).unwrap();

        assert_eq!(sheet.name, "Data");
        assert_eq!(sheet.row_count(), 3);
        assert_eq!(sheet.col_count(), 3);
        assert_eq!(sheet.get_src(&CellRef::new(0, 0)).unwrap(), "name");
        assert_eq!(sheet.get_src(&CellRef::new(2, 0)).unwrap(), "Hello, world");
        assert_eq!(sheet.get_src(&CellRef::new(2, 2)).unwrap(), "line1\nline2");
        assert_eq!(sheet.get_cell_type(&CellRef::new(1, 1)), CellType::Int);
        assert_eq!(sheet.get_cell_type(&CellRef::new(2, 1)), CellType::Bool);
        assert_eq!(sheet.get_result_data(&CellRef::new(1, 2)).to_string(), "24");
    }

    #[test]
    fn accepts_ragged_rows() {
        let sheet = parse_csv_sheet(b"a,b,c\n1\n2,3\n", "Data", None, &Locale::default()).unwrap();

        assert_eq!(sheet.row_count(), 3);
        assert_eq!(sheet.col_count(), 3);
        assert_eq!(sheet.get_src(&CellRef::new(1, 1)).unwrap(), "");
    }

    #[test]
    fn rejects_invalid_utf8() {
        let err = parse_csv_sheet(b"valid,\xff\n", "Data", None, &Locale::default()).unwrap_err();

        assert!(matches!(err, Error::Csv(_)));
    }
}
