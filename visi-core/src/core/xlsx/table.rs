use crate::core::parser::col_idx_to_letters;
use crate::core::table::ExcelTable;
use crate::core::xlsx::xml_util::NamespaceFlavor;

pub(crate) fn normalize_table_style_name(name: &str) -> String {
    let lower = name.trim().to_ascii_lowercase();
    if let Some(rest) = lower.strip_prefix("tablestylelight") {
        return format!("TableStyleLight{rest}");
    }
    if let Some(rest) = lower.strip_prefix("light") {
        return format!("TableStyleLight{rest}");
    }
    if let Some(rest) = lower.strip_prefix("tablestylemedium") {
        return format!("TableStyleMedium{rest}");
    }
    if let Some(rest) = lower.strip_prefix("medium") {
        return format!("TableStyleMedium{rest}");
    }
    if let Some(rest) = lower.strip_prefix("tablestyledark") {
        return format!("TableStyleDark{rest}");
    }
    if let Some(rest) = lower.strip_prefix("dark") {
        return format!("TableStyleDark{rest}");
    }
    if name.trim().is_empty() {
        "TableStyleMedium9".to_string()
    } else {
        name.trim().to_string()
    }
}

pub(crate) fn serialize_table_xml(
    table: &ExcelTable,
    table_id: usize,
    flavor: NamespaceFlavor,
) -> String {
    let ns = flavor.spreadsheetml();
    let is_empty_table = table.has_insert_row
        || (table.has_header_row && table.start_row == table.end_row && !table.has_totals_row);
    let end_row = if is_empty_table {
        table.start_row + 1
    } else {
        table.end_row
    };
    let start_col_let = col_idx_to_letters(table.start_col);
    let end_col_let = col_idx_to_letters(table.end_col);
    let table_ref = format!(
        "{start_col_let}{}:{end_col_let}{}",
        table.start_row + 1,
        end_row + 1
    );
    let escaped_name = super::escape_xml(&table.name);
    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><table xmlns=\"{ns}\" id=\"{table_id}\" name=\"{escaped_name}\" displayName=\"{escaped_name}\" ref=\"{table_ref}\""
    );
    if !table.has_header_row {
        xml.push_str(" headerRowCount=\"0\"");
    }
    if table.has_totals_row {
        xml.push_str(" totalsRowCount=\"1\"");
    } else {
        xml.push_str(" totalsRowShown=\"0\"");
    }
    if is_empty_table {
        xml.push_str(" insertRow=\"1\"");
    }
    xml.push('>');

    if table.has_header_row {
        let af_end_row = if table.has_totals_row {
            end_row.saturating_sub(1).max(table.start_row)
        } else {
            end_row
        };
        let af_ref = format!(
            "{start_col_let}{}:{end_col_let}{}",
            table.start_row + 1,
            af_end_row + 1
        );
        xml.push_str(&format!("<autoFilter ref=\"{af_ref}\"/>"));
    }

    let col_count = table.end_col - table.start_col + 1;
    xml.push_str(&format!("<tableColumns count=\"{col_count}\">"));
    for idx in 0..col_count {
        let col_name = table
            .columns
            .get(idx)
            .filter(|s| !s.is_empty())
            .cloned()
            .unwrap_or_else(|| format!("Column{}", idx + 1));
        xml.push_str(&format!(
            "<tableColumn id=\"{}\" name=\"{}\"/>",
            idx + 1,
            super::escape_xml(&col_name)
        ));
    }
    xml.push_str("</tableColumns>");

    let style_name = table
        .style_name
        .as_deref()
        .map(normalize_table_style_name)
        .unwrap_or_else(|| "TableStyleMedium9".to_string());
    xml.push_str(&format!(
        "<tableStyleInfo name=\"{}\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/>",
        super::escape_xml(&style_name)
    ));
    xml.push_str("</table>");
    xml
}
