use crate::core::chart::{Chart, ChartType};
use crate::core::parser::{Expr, col_idx_to_letters, parse_excel_formula};
use crate::core::xlsx::xml_util::NamespaceFlavor;

fn quote_sheet_name(name: &str) -> String {
    if name.chars().all(|c| c.is_ascii_alphanumeric() || c == '_') {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

fn chart_series_ranges(chart: &Chart, ws_name: &str) -> (Option<String>, String) {
    let sheet_ref = quote_sheet_name(ws_name);
    match parse_excel_formula(&chart.data_range) {
        Ok(Expr::CellRef { row, col, .. }) => {
            let col_letter = col_idx_to_letters(col);
            let values_range = format!(
                "{sheet_ref}!${col_letter}${}:${col_letter}${}",
                row + 1,
                row + 1
            );
            (None, values_range)
        }
        Ok(Expr::RangeRef {
            start_row,
            start_col,
            end_row,
            end_col,
            ..
        }) => {
            let cols_count = end_col + 1 - start_col;
            if cols_count >= 2 {
                let cat_col = col_idx_to_letters(start_col);
                let val_col = col_idx_to_letters(start_col + 1);
                if end_row == usize::MAX {
                    (
                        Some(format!("{sheet_ref}!${cat_col}:${cat_col}")),
                        format!("{sheet_ref}!${val_col}:${val_col}"),
                    )
                } else {
                    (
                        Some(format!(
                            "{sheet_ref}!${cat_col}${}:${cat_col}${}",
                            start_row + 1,
                            end_row + 1
                        )),
                        format!(
                            "{sheet_ref}!${val_col}${}:${val_col}${}",
                            start_row + 1,
                            end_row + 1
                        ),
                    )
                }
            } else {
                let col_letter = col_idx_to_letters(start_col);
                let values_range = if end_row == usize::MAX {
                    format!("{sheet_ref}!${col_letter}:${col_letter}")
                } else {
                    format!(
                        "{sheet_ref}!${col_letter}${}:${col_letter}${}",
                        start_row + 1,
                        end_row + 1
                    )
                };
                (None, values_range)
            }
        }
        _ => (None, chart.data_range.clone()),
    }
}

fn write_title_xml(out: &mut String, text: &str) {
    out.push_str("<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>");
    out.push_str(&super::escape_xml(text));
    out.push_str("</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>");
}

fn write_axis_xml(
    out: &mut String,
    tag: &str,
    ax_id: u32,
    pos: &str,
    cross_ax: u32,
    title: Option<&str>,
) {
    out.push_str(&format!("<c:{tag}><c:axId val=\"{ax_id}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"{pos}\"/>"));
    if let Some(t) = title {
        write_title_xml(out, t);
    }
    out.push_str(&format!(
        "<c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{cross_ax}\"/><c:crosses val=\"autoZero\"/></c:{tag}>"
    ));
}

pub(crate) fn serialize_chart_xml(chart: &Chart, ws_name: &str, flavor: NamespaceFlavor) -> String {
    let ns_c = flavor.drawingml_chart();
    let ns_a = flavor.drawingml_main();
    let ns_r = flavor.office_rels();
    let (cat_range, val_range) = chart_series_ranges(chart, ws_name);

    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><c:chartSpace xmlns:c=\"{ns_c}\" xmlns:a=\"{ns_a}\" xmlns:r=\"{ns_r}\"><c:chart>"
    );

    if let Some(ref title) = chart.title {
        write_title_xml(&mut xml, title);
        xml.push_str("<c:autoTitleDeleted val=\"0\"/>");
    } else {
        xml.push_str("<c:autoTitleDeleted val=\"1\"/>");
    }

    xml.push_str("<c:plotArea><c:layout/>");

    let esc_val = super::escape_xml(&val_range);
    let esc_cat = cat_range.as_deref().map(super::escape_xml);

    match chart.chart_type {
        ChartType::Column | ChartType::Bar => {
            let dir = if chart.chart_type == ChartType::Bar {
                "bar"
            } else {
                "col"
            };
            xml.push_str(&format!(
                "<c:barChart><c:barDir val=\"{dir}\"/><c:grouping val=\"clustered\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
            ));
            if let Some(ref cat) = esc_cat {
                xml.push_str(&format!(
                    "<c:cat><c:strRef><c:f>{cat}</c:f></c:strRef></c:cat>"
                ));
            }
            xml.push_str(&format!(
                "<c:val><c:numRef><c:f>{esc_val}</c:f></c:numRef></c:val></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>"
            ));
        }
        ChartType::Line => {
            xml.push_str(
                "<c:lineChart><c:grouping val=\"standard\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
            );
            if let Some(ref cat) = esc_cat {
                xml.push_str(&format!(
                    "<c:cat><c:strRef><c:f>{cat}</c:f></c:strRef></c:cat>"
                ));
            }
            xml.push_str(&format!(
                "<c:val><c:numRef><c:f>{esc_val}</c:f></c:numRef></c:val></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>"
            ));
        }
        ChartType::Area => {
            xml.push_str(
                "<c:areaChart><c:grouping val=\"standard\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
            );
            if let Some(ref cat) = esc_cat {
                xml.push_str(&format!(
                    "<c:cat><c:strRef><c:f>{cat}</c:f></c:strRef></c:cat>"
                ));
            }
            xml.push_str(&format!(
                "<c:val><c:numRef><c:f>{esc_val}</c:f></c:numRef></c:val></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:areaChart>"
            ));
        }
        ChartType::Pie => {
            xml.push_str(
                "<c:pieChart><c:varyColors val=\"1\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
            );
            if let Some(ref cat) = esc_cat {
                xml.push_str(&format!(
                    "<c:cat><c:strRef><c:f>{cat}</c:f></c:strRef></c:cat>"
                ));
            }
            xml.push_str(&format!(
                "<c:val><c:numRef><c:f>{esc_val}</c:f></c:numRef></c:val></c:ser></c:pieChart>"
            ));
        }
        ChartType::Scatter => {
            xml.push_str(
                "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
            );
            if let Some(ref cat) = esc_cat {
                xml.push_str(&format!(
                    "<c:xVal><c:numRef><c:f>{cat}</c:f></c:numRef></c:xVal>"
                ));
            }
            xml.push_str(&format!(
                "<c:yVal><c:numRef><c:f>{esc_val}</c:f></c:numRef></c:yVal></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:scatterChart>"
            ));
        }
    }

    match chart.chart_type {
        ChartType::Pie => {}
        ChartType::Scatter => {
            write_axis_xml(&mut xml, "valAx", 1, "b", 2, chart.xlabel.as_deref());
            write_axis_xml(&mut xml, "valAx", 2, "l", 1, chart.ylabel.as_deref());
        }
        ChartType::Bar => {
            write_axis_xml(&mut xml, "catAx", 1, "l", 2, chart.xlabel.as_deref());
            write_axis_xml(&mut xml, "valAx", 2, "b", 1, chart.ylabel.as_deref());
        }
        ChartType::Column | ChartType::Line | ChartType::Area => {
            write_axis_xml(&mut xml, "catAx", 1, "b", 2, chart.xlabel.as_deref());
            write_axis_xml(&mut xml, "valAx", 2, "l", 1, chart.ylabel.as_deref());
        }
    }

    xml.push_str("</c:plotArea>");

    if chart.show_legend {
        xml.push_str("<c:legend><c:legendPos val=\"r\"/><c:overlay val=\"0\"/></c:legend>");
    }

    xml.push_str("<c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>");
    xml
}

pub(crate) fn serialize_drawing_xml(
    anchors: &[(usize, usize, String)],
    flavor: NamespaceFlavor,
) -> String {
    let ns_xdr = flavor.drawingml_spreadsheet();
    let ns_a = flavor.drawingml_main();
    let ns_c = flavor.drawingml_chart();
    let ns_r = flavor.office_rels();

    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"{ns_xdr}\" xmlns:a=\"{ns_a}\">"
    );
    for (idx, (anchor_row, anchor_col, r_id)) in anchors.iter().enumerate() {
        let to_col = anchor_col + 8;
        let to_row = anchor_row + 15;
        let obj_id = idx + 2;
        xml.push_str(&format!(
            "<xdr:twoCellAnchor><xdr:from><xdr:col>{anchor_col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{anchor_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{to_col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{to_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{obj_id}\" name=\"Chart {}\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"{ns_c}\"><c:chart xmlns:c=\"{ns_c}\" xmlns:r=\"{ns_r}\" r:id=\"{r_id}\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>",
            idx + 1
        ));
    }
    xml.push_str("</xdr:wsDr>");
    xml
}
