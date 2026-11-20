use std::collections::HashMap;

use quick_xml::events::Event;

use crate::core::engine::CellComment;
use crate::core::parser::col_idx_to_letters;
use crate::core::xlsx::xml_util::{
    NamespaceFlavor, decode_xstring, encode_xstring, needs_xml_space_preserve,
};

pub(crate) fn parse_comments_xml(xml: &str) -> Vec<CellComment> {
    let mut authors: Vec<String> = Vec::new();
    let mut comments: Vec<CellComment> = Vec::new();

    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();

    let mut in_author = false;
    let mut author_buf = String::new();

    let mut in_comment = false;
    let mut comment_rc: Option<(usize, usize)> = None;
    let mut comment_author_id: usize = 0;
    let mut comment_text = String::new();
    let mut in_t = false;
    let mut phonetic_depth = 0usize;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"author" => {
                    in_author = true;
                    author_buf.clear();
                }
                b"comment" => {
                    in_comment = true;
                    comment_rc = super::get_attr(e, b"ref").and_then(|r| super::parse_a1_cell(&r));
                    comment_author_id = super::get_attr(e, b"authorId")
                        .and_then(|s| s.parse::<usize>().ok())
                        .unwrap_or(0);
                    comment_text.clear();
                }
                b"rPh" if in_comment => {
                    phonetic_depth += 1;
                }
                b"t" if in_comment && phonetic_depth == 0 => {
                    in_t = true;
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"author" {
                    authors.push(String::new());
                }
            }
            Ok(Event::Text(e)) => {
                if let Ok(decoded) = e.decode()
                    && let Ok(unescaped) = quick_xml::escape::unescape(&decoded)
                {
                    if in_author {
                        author_buf.push_str(&unescaped);
                    } else if in_t && phonetic_depth == 0 {
                        comment_text.push_str(&unescaped);
                    }
                }
            }
            Ok(Event::GeneralRef(e)) => {
                if let Ok(raw) = e.decode() {
                    let entity = format!("&{raw};");
                    let resolved = quick_xml::escape::unescape(&entity)
                        .map(|v| v.into_owned())
                        .unwrap_or(entity);
                    if in_author {
                        author_buf.push_str(&resolved);
                    } else if in_t && phonetic_depth == 0 {
                        comment_text.push_str(&resolved);
                    }
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"author" => {
                    in_author = false;
                    authors.push(decode_xstring(&author_buf));
                    author_buf.clear();
                }
                b"rPh" if phonetic_depth > 0 => {
                    phonetic_depth -= 1;
                }
                b"t" if in_comment => {
                    in_t = false;
                }
                b"comment" => {
                    in_comment = false;
                    in_t = false;
                    phonetic_depth = 0;
                    if let Some((row, col)) = comment_rc.take() {
                        let author = authors.get(comment_author_id).cloned().unwrap_or_default();
                        comments.push(CellComment {
                            row,
                            col,
                            author,
                            text: decode_xstring(&comment_text),
                        });
                    }
                    comment_text.clear();
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    comments
}

pub(crate) fn write_comments_xml(comments: &[CellComment], flavor: NamespaceFlavor) -> String {
    let ns = flavor.spreadsheetml();
    let mut authors: Vec<String> = Vec::new();
    let mut author_idx: HashMap<String, usize> = HashMap::new();

    for c in comments {
        if !author_idx.contains_key(&c.author) {
            let idx = authors.len();
            authors.push(c.author.clone());
            author_idx.insert(c.author.clone(), idx);
        }
    }

    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><comments xmlns=\"{ns}\"><authors>"
    );
    for author in &authors {
        xml.push_str("<author>");
        xml.push_str(&super::escape_xml_text(&encode_xstring(author)));
        xml.push_str("</author>");
    }
    xml.push_str("</authors><commentList>");
    for c in comments {
        let cell_ref = format!("{}{}", col_idx_to_letters(c.col), c.row + 1);
        let aid = author_idx.get(&c.author).copied().unwrap_or(0);
        let encoded = encode_xstring(&c.text);
        let space_attr = if needs_xml_space_preserve(&encoded) {
            " xml:space=\"preserve\""
        } else {
            ""
        };
        xml.push_str(&format!(
            "<comment ref=\"{cell_ref}\" authorId=\"{aid}\"><text><t{space_attr}>{}</t></text></comment>",
            super::escape_xml_text(&encoded)
        ));
    }
    xml.push_str("</commentList></comments>");
    xml
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn comments_round_trip() {
        let input = vec![
            CellComment {
                row: 0,
                col: 1,
                author: "Alice".to_string(),
                text: "Check Q1 total".to_string(),
            },
            CellComment {
                row: 2,
                col: 0,
                author: "Bob".to_string(),
                text: " line1\nline2 ".to_string(),
            },
        ];
        let xml = write_comments_xml(&input, NamespaceFlavor::Transitional);
        let parsed = parse_comments_xml(&xml);
        assert_eq!(parsed, input);
    }
}
