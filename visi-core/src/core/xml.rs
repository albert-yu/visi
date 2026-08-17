// FILTERXML support: a minimal DOM built with quick_xml plus a hand-rolled
// evaluator for the small XPath subset Excel's own FILTERXML documents
// (absolute child/descendant paths, `*`, `@attr`, `text()`, `[n]` and
// `[@attr='val']` predicates). Full XPath 1.0 is out of scope -- Excel
// itself rejects anything outside this subset.

#[derive(Debug, Clone)]
struct XmlNode {
    name: String,
    attrs: Vec<(String, String)>,
    children: Vec<XmlNode>,
    text: String,
}

fn local_name(e: &quick_xml::events::BytesStart) -> String {
    String::from_utf8_lossy(e.name().local_name().into_inner()).to_string()
}

fn read_attrs(e: &quick_xml::events::BytesStart) -> Result<Vec<(String, String)>, String> {
    let mut attrs = Vec::new();
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let local = match key.iter().position(|&b| b == b':') {
            Some(pos) => &key[pos + 1..],
            None => key,
        };
        let key = String::from_utf8_lossy(local).to_string();
        let value = attr
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, e.decoder())
            .map_err(|_| "#VALUE!".to_string())?
            .to_string();
        attrs.push((key, value));
    }
    Ok(attrs)
}

fn parse_xml(xml: &str) -> Result<XmlNode, String> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut stack: Vec<XmlNode> = Vec::new();
    let mut root: Option<XmlNode> = None;

    loop {
        let event = match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => return Err("#VALUE!".to_string()),
            Ok(e) => e,
        };
        match event {
            quick_xml::events::Event::Start(ref e) => {
                stack.push(XmlNode {
                    name: local_name(e),
                    attrs: read_attrs(e)?,
                    children: Vec::new(),
                    text: String::new(),
                });
            }
            quick_xml::events::Event::Empty(ref e) => {
                let node = XmlNode {
                    name: local_name(e),
                    attrs: read_attrs(e)?,
                    children: Vec::new(),
                    text: String::new(),
                };
                match stack.last_mut() {
                    Some(parent) => parent.children.push(node),
                    None => root = Some(node),
                }
            }
            quick_xml::events::Event::Text(e) => {
                let decoded = e.decode().map_err(|_| "#VALUE!".to_string())?;
                let txt = quick_xml::escape::unescape(&decoded)
                    .map_err(|_| "#VALUE!".to_string())?
                    .to_string();
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&txt);
                }
            }
            quick_xml::events::Event::End(_) => {
                if let Some(node) = stack.pop() {
                    match stack.last_mut() {
                        Some(parent) => parent.children.push(node),
                        None => root = Some(node),
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }

    root.ok_or_else(|| "#VALUE!".to_string())
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum Axis {
    Child,
    Descendant,
}

#[derive(Debug, Clone, PartialEq)]
enum Test {
    Name(String),
    Wildcard,
    Attr(String),
    Text,
}

#[derive(Debug, Clone)]
enum Predicate {
    Index(usize),
    AttrEq(String, String),
}

#[derive(Debug, Clone)]
struct Step {
    axis: Axis,
    test: Test,
    predicate: Option<Predicate>,
}

fn parse_step_with_axis(seg: &str, axis: Axis) -> Result<Step, String> {
    let (test_str, pred_str) = match seg.find('[') {
        Some(idx) => {
            if !seg.ends_with(']') {
                return Err("#VALUE!".to_string());
            }
            (&seg[..idx], Some(&seg[idx + 1..seg.len() - 1]))
        }
        None => (seg, None),
    };
    if test_str.is_empty() {
        return Err("#VALUE!".to_string());
    }
    let test = if test_str == "*" {
        Test::Wildcard
    } else if test_str == "text()" {
        Test::Text
    } else if let Some(attr) = test_str.strip_prefix('@') {
        Test::Attr(attr.to_string())
    } else {
        Test::Name(test_str.to_string())
    };
    let predicate = match pred_str {
        None => None,
        Some(p) => {
            let p = p.trim();
            if let Ok(n) = p.parse::<usize>() {
                Some(Predicate::Index(n))
            } else if let Some(rest) = p.strip_prefix('@') {
                let eq = rest.find('=').ok_or_else(|| "#VALUE!".to_string())?;
                let name = rest[..eq].trim().to_string();
                let val = rest[eq + 1..]
                    .trim()
                    .trim_matches(|c| c == '\'' || c == '"')
                    .to_string();
                Some(Predicate::AttrEq(name, val))
            } else {
                return Err("#VALUE!".to_string());
            }
        }
    };
    Ok(Step {
        axis,
        test,
        predicate,
    })
}

fn parse_path(path: &str) -> Result<Vec<Step>, String> {
    let path = path.trim();
    let mut chars = path.chars().peekable();
    let mut axis = Axis::Child;
    match chars.peek() {
        Some('/') => {
            chars.next();
            if chars.peek() == Some(&'/') {
                chars.next();
                axis = Axis::Descendant;
            }
        }
        _ => return Err("#VALUE!".to_string()),
    }

    let mut steps = Vec::new();
    loop {
        let mut seg = String::new();
        while let Some(&c) = chars.peek() {
            if c == '/' {
                break;
            }
            seg.push(c);
            chars.next();
        }
        if seg.is_empty() {
            return Err("#VALUE!".to_string());
        }
        steps.push(parse_step_with_axis(&seg, axis)?);

        if chars.peek().is_none() {
            break;
        }
        chars.next(); // consume '/'
        axis = if chars.peek() == Some(&'/') {
            chars.next();
            Axis::Descendant
        } else {
            Axis::Child
        };
    }
    Ok(steps)
}

fn matches_test(node: &XmlNode, test: &Test) -> bool {
    match test {
        Test::Wildcard => true,
        Test::Name(name) => &node.name == name,
        Test::Attr(_) | Test::Text => false,
    }
}

fn apply_predicate<'a>(nodes: Vec<&'a XmlNode>, predicate: &Option<Predicate>) -> Vec<&'a XmlNode> {
    match predicate {
        None => nodes,
        Some(Predicate::Index(n)) => {
            if *n >= 1 && *n <= nodes.len() {
                vec![nodes[*n - 1]]
            } else {
                vec![]
            }
        }
        Some(Predicate::AttrEq(key, val)) => nodes
            .into_iter()
            .filter(|n| n.attrs.iter().any(|(k, v)| k == key && v == val))
            .collect(),
    }
}

fn collect_descendants<'a>(node: &'a XmlNode, test: &Test, out: &mut Vec<&'a XmlNode>) {
    for child in &node.children {
        if matches_test(child, test) {
            out.push(child);
        }
        collect_descendants(child, test, out);
    }
}

fn step_nodes<'a>(input: &[&'a XmlNode], step: &Step, is_first: bool) -> Vec<&'a XmlNode> {
    if is_first {
        let matched: Vec<&XmlNode> = match step.axis {
            Axis::Child => input
                .iter()
                .filter(|n| matches_test(n, &step.test))
                .copied()
                .collect(),
            // A leading `//` searches the whole document, including the
            // root element itself, not just its children.
            Axis::Descendant => {
                let mut group = Vec::new();
                for node in input {
                    if matches_test(node, &step.test) {
                        group.push(*node);
                    }
                    collect_descendants(node, &step.test, &mut group);
                }
                group
            }
        };
        return apply_predicate(matched, &step.predicate);
    }
    let mut result = Vec::new();
    for node in input {
        let group: Vec<&XmlNode> = match step.axis {
            Axis::Child => node
                .children
                .iter()
                .filter(|c| matches_test(c, &step.test))
                .collect(),
            Axis::Descendant => {
                let mut out = Vec::new();
                collect_descendants(node, &step.test, &mut out);
                out
            }
        };
        result.extend(apply_predicate(group, &step.predicate));
    }
    result
}

fn evaluate(root: &XmlNode, xpath: &str) -> Result<String, String> {
    let steps = parse_path(xpath)?;
    if steps.is_empty() {
        return Err("#VALUE!".to_string());
    }

    let mut nodes: Vec<&XmlNode> = vec![root];
    for (i, step) in steps[..steps.len() - 1].iter().enumerate() {
        if matches!(step.test, Test::Attr(_) | Test::Text) {
            // Only supported as the final step, matching the common
            // FILTERXML usage of `@attr`/`text()` at the end of a path.
            return Err("#VALUE!".to_string());
        }
        nodes = step_nodes(&nodes, step, i == 0);
        if nodes.is_empty() {
            return Err("#N/A".to_string());
        }
    }

    let last = steps.last().unwrap();
    let is_first = steps.len() == 1;
    match &last.test {
        Test::Attr(name) => {
            let candidates = if is_first {
                apply_predicate(nodes.clone(), &last.predicate)
            } else {
                nodes.clone()
            };
            for node in &candidates {
                if let Some((_, v)) = node.attrs.iter().find(|(k, _)| k == name) {
                    return Ok(v.clone());
                }
            }
            Err("#N/A".to_string())
        }
        Test::Text => {
            let candidates = if is_first {
                apply_predicate(nodes.clone(), &last.predicate)
            } else {
                nodes.clone()
            };
            match candidates.first() {
                Some(node) => Ok(node.text.clone()),
                None => Err("#N/A".to_string()),
            }
        }
        _ => {
            let matched = step_nodes(&nodes, last, is_first);
            match matched.first() {
                Some(node) => Ok(node.text.clone()),
                None => Err("#N/A".to_string()),
            }
        }
    }
}

pub fn filterxml(xml: &str, xpath: &str) -> Result<String, String> {
    let root = parse_xml(xml)?;
    evaluate(&root, xpath)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn simple_element_text() {
        let xml = "<a><b>hello</b></a>";
        assert_eq!(filterxml(xml, "/a/b"), Ok("hello".to_string()));
    }

    #[test]
    fn attribute_value() {
        let xml = "<a><b id=\"5\">hi</b></a>";
        assert_eq!(filterxml(xml, "/a/b/@id"), Ok("5".to_string()));
    }

    #[test]
    fn text_function() {
        let xml = "<a><b>hi there</b></a>";
        assert_eq!(filterxml(xml, "/a/b/text()"), Ok("hi there".to_string()));
    }

    #[test]
    fn descendant_axis() {
        let xml = "<a><c><b>deep</b></c></a>";
        assert_eq!(filterxml(xml, "//b"), Ok("deep".to_string()));
    }

    #[test]
    fn attribute_predicate() {
        let xml = "<a><b id=\"1\">one</b><b id=\"2\">two</b></a>";
        assert_eq!(filterxml(xml, "/a/b[@id='2']"), Ok("two".to_string()));
    }

    #[test]
    fn index_predicate() {
        let xml = "<a><b>one</b><b>two</b></a>";
        assert_eq!(filterxml(xml, "/a/b[2]"), Ok("two".to_string()));
    }

    #[test]
    fn no_match_is_na() {
        let xml = "<a><b>one</b></a>";
        assert_eq!(filterxml(xml, "/a/c"), Err("#N/A".to_string()));
    }

    #[test]
    fn malformed_xml_is_value_error() {
        assert_eq!(filterxml("<a><b>", "/a/b"), Err("#VALUE!".to_string()));
    }
}
