use std::collections::HashMap;
use std::rc::Rc;

use regex::RegexBuilder;

use super::host::ObjRef;
use super::value::{self, VResult, VarArray, Variant, VbaError};

#[derive(Debug, Clone, Default)]
pub struct VbaCollection {
    entries: Vec<(Option<String>, Variant)>,
}

impl VbaCollection {
    pub fn count(&self) -> i32 {
        self.entries.len() as i32
    }

    pub fn add(
        &mut self,
        item: Variant,
        key: Option<String>,
        before: Option<&Variant>,
        after: Option<&Variant>,
    ) -> VResult<()> {
        if before.is_some() && after.is_some() {
            return Err(VbaError::invalid_call());
        }
        if let Some(ref k) = key
            && self
                .entries
                .iter()
                .any(|(ek, _)| ek.as_ref().is_some_and(|e| e.eq_ignore_ascii_case(k)))
        {
            return Err(VbaError::new(
                457,
                "This key is already associated with an element of this collection",
            ));
        }
        let pos = if let Some(b) = before {
            self.resolve_index(b)?
        } else if let Some(a) = after {
            self.resolve_index(a)? + 1
        } else {
            self.entries.len()
        };
        self.entries.insert(pos, (key, item));
        Ok(())
    }

    pub fn item(&self, index: &Variant) -> VResult<Variant> {
        let idx = self.resolve_index(index)?;
        Ok(self.entries[idx].1.clone())
    }

    pub fn remove(&mut self, index: &Variant) -> VResult<()> {
        let idx = self.resolve_index(index)?;
        self.entries.remove(idx);
        Ok(())
    }

    pub fn items(&self) -> Vec<Variant> {
        self.entries.iter().map(|(_, v)| v.clone()).collect()
    }

    fn resolve_index(&self, index: &Variant) -> VResult<usize> {
        if let Variant::Str(k) = index {
            return self
                .entries
                .iter()
                .position(|(ek, _)| ek.as_ref().is_some_and(|e| e.eq_ignore_ascii_case(k)))
                .ok_or_else(VbaError::invalid_call);
        }
        let n = value::bankers_round(index.to_f64()?) as i64;
        if n < 1 || (n as usize) > self.entries.len() {
            return Err(VbaError::subscript());
        }
        Ok(n as usize - 1)
    }
}

#[derive(Debug, Clone, Default)]
pub struct ScriptingDictionary {
    entries: Vec<(Variant, Variant)>,
    pub compare_mode: i32,
}

impl ScriptingDictionary {
    fn key_eq(&self, a: &Variant, b: &Variant) -> bool {
        if self.compare_mode == 1
            && let (Variant::Str(sa), Variant::Str(sb)) = (a, b)
        {
            return sa.to_lowercase() == sb.to_lowercase();
        }
        a == b
    }

    fn find_index(&self, key: &Variant) -> Option<usize> {
        self.entries.iter().position(|(k, _)| self.key_eq(k, key))
    }

    pub fn count(&self) -> i32 {
        self.entries.len() as i32
    }

    pub fn exists(&self, key: &Variant) -> bool {
        self.find_index(key).is_some()
    }

    pub fn add(&mut self, key: Variant, item: Variant) -> VResult<()> {
        if self.find_index(&key).is_some() {
            return Err(VbaError::new(
                457,
                "This key is already associated with an element of this collection",
            ));
        }
        self.entries.push((key, item));
        Ok(())
    }

    pub fn get_item(&mut self, key: &Variant) -> Variant {
        if let Some(idx) = self.find_index(key) {
            self.entries[idx].1.clone()
        } else {
            self.entries.push((key.clone(), Variant::Empty));
            Variant::Empty
        }
    }

    pub fn set_item(&mut self, key: Variant, item: Variant) {
        if let Some(idx) = self.find_index(&key) {
            self.entries[idx].1 = item;
        } else {
            self.entries.push((key, item));
        }
    }

    pub fn rename_key(&mut self, old_key: &Variant, new_key: Variant) -> VResult<()> {
        let Some(idx) = self.find_index(old_key) else {
            return Err(VbaError::new(32811, "Element not found"));
        };
        if self.find_index(&new_key).is_some_and(|j| j != idx) {
            return Err(VbaError::new(
                457,
                "This key is already associated with an element of this collection",
            ));
        }
        self.entries[idx].0 = new_key;
        Ok(())
    }

    pub fn remove(&mut self, key: &Variant) -> VResult<()> {
        let Some(idx) = self.find_index(key) else {
            return Err(VbaError::new(32811, "Element not found"));
        };
        self.entries.remove(idx);
        Ok(())
    }

    pub fn remove_all(&mut self) {
        self.entries.clear();
    }

    pub fn keys(&self) -> Vec<Variant> {
        self.entries.iter().map(|(k, _)| k.clone()).collect()
    }

    pub fn items(&self) -> Vec<Variant> {
        self.entries.iter().map(|(_, v)| v.clone()).collect()
    }
}

#[derive(Debug, Clone, Default)]
pub struct VbScriptRegExp {
    pub pattern: String,
    pub global: bool,
    pub ignore_case: bool,
    pub multi_line: bool,
}

#[derive(Debug, Clone)]
pub struct RegExpMatchData {
    pub value: String,
    pub first_index: i32,
    pub length: i32,
    pub sub_matches: Vec<Variant>,
}

impl VbScriptRegExp {
    fn compile(&self) -> VResult<regex::Regex> {
        RegexBuilder::new(&self.pattern)
            .case_insensitive(self.ignore_case)
            .multi_line(self.multi_line)
            .build()
            .map_err(|e| VbaError::new(5017, format!("Syntax error in regular expression: {e}")))
    }

    pub fn test(&self, input: &str) -> VResult<bool> {
        Ok(self.compile()?.is_match(input))
    }

    pub fn replace(&self, input: &str, replacement: &str) -> VResult<String> {
        let re = self.compile()?;
        if self.global {
            Ok(re.replace_all(input, replacement).into_owned())
        } else {
            Ok(re.replace(input, replacement).into_owned())
        }
    }

    pub fn execute(&self, input: &str) -> VResult<Vec<RegExpMatchData>> {
        let re = self.compile()?;
        let mut out = Vec::new();
        for caps in re.captures_iter(input) {
            if let Some(m) = caps.get(0) {
                let first_index = input[..m.start()].chars().count() as i32;
                let value = m.as_str().to_string();
                let length = value.chars().count() as i32;
                let mut sub_matches = Vec::new();
                for i in 1..caps.len() {
                    match caps.get(i) {
                        Some(sm) => sub_matches.push(Variant::Str(sm.as_str().to_string())),
                        None => sub_matches.push(Variant::Empty),
                    }
                }
                out.push(RegExpMatchData {
                    value,
                    first_index,
                    length,
                    sub_matches,
                });
            }
            if !self.global {
                break;
            }
        }
        Ok(out)
    }
}

#[derive(Debug, Clone, Default)]
pub struct ComStore {
    next_id: u64,
    collections: HashMap<u64, VbaCollection>,
    dictionaries: HashMap<u64, ScriptingDictionary>,
    regexps: HashMap<u64, VbScriptRegExp>,
    match_collections: HashMap<u64, Vec<RegExpMatchData>>,
}

impl ComStore {
    fn alloc_id(&mut self) -> u64 {
        self.next_id += 1;
        self.next_id
    }

    pub fn new_collection(&mut self) -> ObjRef {
        let id = self.alloc_id();
        self.collections.insert(id, VbaCollection::default());
        ObjRef::Collection(id)
    }

    pub fn new_dictionary(&mut self) -> ObjRef {
        let id = self.alloc_id();
        self.dictionaries.insert(id, ScriptingDictionary::default());
        ObjRef::Dictionary(id)
    }

    pub fn new_regexp(&mut self) -> ObjRef {
        let id = self.alloc_id();
        self.regexps.insert(id, VbScriptRegExp::default());
        ObjRef::RegExp(id)
    }

    pub fn create_object(&mut self, prog_id: &str) -> VResult<ObjRef> {
        match prog_id.trim().to_ascii_lowercase().as_str() {
            "scripting.dictionary" => Ok(self.new_dictionary()),
            "vbscript.regexp" => Ok(self.new_regexp()),
            _ => Err(VbaError::new(
                429,
                format!("ActiveX component can't create object: {prog_id}"),
            )),
        }
    }

    pub fn get_member(&mut self, obj: &ObjRef, name: &str, args: &[Variant]) -> VResult<Variant> {
        let lower = name.to_ascii_lowercase();
        match obj {
            ObjRef::Collection(id) => {
                let col = self
                    .collections
                    .get_mut(id)
                    .ok_or_else(|| VbaError::new(91, "Object variable not set"))?;
                match lower.as_str() {
                    "count" => Ok(Variant::Long(col.count())),
                    "item" => {
                        let idx = args.first().ok_or_else(VbaError::invalid_call)?;
                        col.item(idx)
                    }
                    "add" => {
                        let item = args.first().cloned().ok_or_else(VbaError::invalid_call)?;
                        let key = match args.get(1) {
                            Some(v) if !v.is_empty() && !matches!(v, Variant::Missing) => {
                                Some(v.to_vba_string()?)
                            }
                            _ => None,
                        };
                        let before = args
                            .get(2)
                            .filter(|v| !v.is_empty() && !matches!(v, Variant::Missing));
                        let after = args
                            .get(3)
                            .filter(|v| !v.is_empty() && !matches!(v, Variant::Missing));
                        col.add(item, key, before, after)?;
                        Ok(Variant::Empty)
                    }
                    "remove" => {
                        let idx = args.first().ok_or_else(VbaError::invalid_call)?;
                        col.remove(idx)?;
                        Ok(Variant::Empty)
                    }
                    _ => Err(VbaError::new(
                        438,
                        format!(
                            "Object doesn't support this property or method: Collection.{name}"
                        ),
                    )),
                }
            }
            ObjRef::Dictionary(id) => {
                let dict = self
                    .dictionaries
                    .get_mut(id)
                    .ok_or_else(|| VbaError::new(91, "Object variable not set"))?;
                match lower.as_str() {
                    "count" => Ok(Variant::Long(dict.count())),
                    "comparemode" => Ok(Variant::Long(dict.compare_mode)),
                    "exists" => {
                        let key = args.first().ok_or_else(VbaError::invalid_call)?;
                        Ok(Variant::Boolean(dict.exists(key)))
                    }
                    "item" => {
                        let key = args.first().ok_or_else(VbaError::invalid_call)?;
                        Ok(dict.get_item(key))
                    }
                    "add" => {
                        let key = args.first().cloned().ok_or_else(VbaError::invalid_call)?;
                        let item = args.get(1).cloned().ok_or_else(VbaError::invalid_call)?;
                        dict.add(key, item)?;
                        Ok(Variant::Empty)
                    }
                    "remove" => {
                        let key = args.first().ok_or_else(VbaError::invalid_call)?;
                        dict.remove(key)?;
                        Ok(Variant::Empty)
                    }
                    "removeall" => {
                        dict.remove_all();
                        Ok(Variant::Empty)
                    }
                    "keys" => Ok(Variant::Array(Rc::new(VarArray::new_1d(0, dict.keys())))),
                    "items" => Ok(Variant::Array(Rc::new(VarArray::new_1d(0, dict.items())))),
                    _ => Err(VbaError::new(
                        438,
                        format!(
                            "Object doesn't support this property or method: Dictionary.{name}"
                        ),
                    )),
                }
            }
            ObjRef::RegExp(id) => {
                let re = self
                    .regexps
                    .get(id)
                    .cloned()
                    .ok_or_else(|| VbaError::new(91, "Object variable not set"))?;
                match lower.as_str() {
                    "pattern" => Ok(Variant::Str(re.pattern)),
                    "global" => Ok(Variant::Boolean(re.global)),
                    "ignorecase" => Ok(Variant::Boolean(re.ignore_case)),
                    "multiline" => Ok(Variant::Boolean(re.multi_line)),
                    "test" => {
                        let s = args
                            .first()
                            .ok_or_else(VbaError::invalid_call)?
                            .to_vba_string()?;
                        Ok(Variant::Boolean(re.test(&s)?))
                    }
                    "replace" => {
                        let s = args
                            .first()
                            .ok_or_else(VbaError::invalid_call)?
                            .to_vba_string()?;
                        let repl = args
                            .get(1)
                            .ok_or_else(VbaError::invalid_call)?
                            .to_vba_string()?;
                        Ok(Variant::Str(re.replace(&s, &repl)?))
                    }
                    "execute" => {
                        let s = args
                            .first()
                            .ok_or_else(VbaError::invalid_call)?
                            .to_vba_string()?;
                        let matches = re.execute(&s)?;
                        let mc_id = self.alloc_id();
                        self.match_collections.insert(mc_id, matches);
                        Ok(Variant::Object(ObjRef::RegExpMatches(mc_id)))
                    }
                    _ => Err(VbaError::new(
                        438,
                        format!("Object doesn't support this property or method: RegExp.{name}"),
                    )),
                }
            }
            ObjRef::RegExpMatches(id) => {
                let mc = self
                    .match_collections
                    .get(id)
                    .ok_or_else(|| VbaError::new(91, "Object variable not set"))?;
                match lower.as_str() {
                    "count" => Ok(Variant::Long(mc.len() as i32)),
                    "item" => {
                        let idx = value::bankers_round(
                            args.first().ok_or_else(VbaError::invalid_call)?.to_f64()?,
                        ) as i64;
                        if idx < 0 || (idx as usize) >= mc.len() {
                            return Err(VbaError::subscript());
                        }
                        Ok(Variant::Object(ObjRef::RegExpMatch(*id, idx as usize)))
                    }
                    _ => Err(VbaError::new(
                        438,
                        format!(
                            "Object doesn't support this property or method: MatchCollection.{name}"
                        ),
                    )),
                }
            }
            ObjRef::RegExpMatch(mc_id, idx) => {
                let m = self
                    .match_collections
                    .get(mc_id)
                    .and_then(|mc| mc.get(*idx))
                    .ok_or_else(|| VbaError::new(91, "Object variable not set"))?;
                match lower.as_str() {
                    "value" => Ok(Variant::Str(m.value.clone())),
                    "firstindex" => Ok(Variant::Long(m.first_index)),
                    "length" => Ok(Variant::Long(m.length)),
                    "submatches" => {
                        if args.is_empty() {
                            Ok(Variant::Object(ObjRef::RegExpSubMatches(*mc_id, *idx)))
                        } else {
                            let sidx = value::bankers_round(args[0].to_f64()?) as i64;
                            if sidx < 0 || (sidx as usize) >= m.sub_matches.len() {
                                return Err(VbaError::subscript());
                            }
                            Ok(m.sub_matches[sidx as usize].clone())
                        }
                    }
                    _ => Err(VbaError::new(
                        438,
                        format!("Object doesn't support this property or method: Match.{name}"),
                    )),
                }
            }
            ObjRef::RegExpSubMatches(mc_id, idx) => {
                let m = self
                    .match_collections
                    .get(mc_id)
                    .and_then(|mc| mc.get(*idx))
                    .ok_or_else(|| VbaError::new(91, "Object variable not set"))?;
                match lower.as_str() {
                    "count" => Ok(Variant::Long(m.sub_matches.len() as i32)),
                    "item" => {
                        let sidx = value::bankers_round(
                            args.first().ok_or_else(VbaError::invalid_call)?.to_f64()?,
                        ) as i64;
                        if sidx < 0 || (sidx as usize) >= m.sub_matches.len() {
                            return Err(VbaError::subscript());
                        }
                        Ok(m.sub_matches[sidx as usize].clone())
                    }
                    _ => Err(VbaError::new(
                        438,
                        format!(
                            "Object doesn't support this property or method: SubMatches.{name}"
                        ),
                    )),
                }
            }
            _ => Err(VbaError::new(
                438,
                format!("Object doesn't support this property or method: .{name}"),
            )),
        }
    }

    pub fn set_member(
        &mut self,
        obj: &ObjRef,
        name: &str,
        args: &[Variant],
        value: &Variant,
    ) -> VResult<()> {
        let lower = name.to_ascii_lowercase();
        match obj {
            ObjRef::Dictionary(id) => {
                let dict = self
                    .dictionaries
                    .get_mut(id)
                    .ok_or_else(|| VbaError::new(91, "Object variable not set"))?;
                match lower.as_str() {
                    "comparemode" => {
                        if dict.count() > 0 {
                            return Err(VbaError::invalid_call());
                        }
                        dict.compare_mode = value::bankers_round(value.to_f64()?) as i32;
                        Ok(())
                    }
                    "item" => {
                        let key = args.first().cloned().ok_or_else(VbaError::invalid_call)?;
                        dict.set_item(key, value.clone());
                        Ok(())
                    }
                    "key" => {
                        let old_key = args.first().ok_or_else(VbaError::invalid_call)?;
                        dict.rename_key(old_key, value.clone())
                    }
                    _ => Err(VbaError::new(
                        438,
                        format!(
                            "Object doesn't support this property or method: Dictionary.{name}"
                        ),
                    )),
                }
            }
            ObjRef::RegExp(id) => {
                let re = self
                    .regexps
                    .get_mut(id)
                    .ok_or_else(|| VbaError::new(91, "Object variable not set"))?;
                match lower.as_str() {
                    "pattern" => {
                        re.pattern = value.to_vba_string()?;
                        Ok(())
                    }
                    "global" => {
                        re.global = value.to_bool()?;
                        Ok(())
                    }
                    "ignorecase" => {
                        re.ignore_case = value.to_bool()?;
                        Ok(())
                    }
                    "multiline" => {
                        re.multi_line = value.to_bool()?;
                        Ok(())
                    }
                    _ => Err(VbaError::new(
                        438,
                        format!("Object doesn't support this property or method: RegExp.{name}"),
                    )),
                }
            }
            _ => Err(VbaError::new(
                438,
                format!("Object doesn't support this property or method: .{name} ="),
            )),
        }
    }

    pub fn iterate(&self, obj: &ObjRef) -> Option<VResult<Vec<Variant>>> {
        match obj {
            ObjRef::Collection(id) => Some(
                self.collections
                    .get(id)
                    .map(|c| c.items())
                    .ok_or_else(|| VbaError::new(91, "Object variable not set")),
            ),
            ObjRef::Dictionary(id) => Some(
                self.dictionaries
                    .get(id)
                    .map(|d| d.keys())
                    .ok_or_else(|| VbaError::new(91, "Object variable not set")),
            ),
            ObjRef::RegExpMatches(id) => Some(
                self.match_collections
                    .get(id)
                    .map(|mc| {
                        (0..mc.len())
                            .map(|i| Variant::Object(ObjRef::RegExpMatch(*id, i)))
                            .collect()
                    })
                    .ok_or_else(|| VbaError::new(91, "Object variable not set")),
            ),
            ObjRef::RegExpSubMatches(mc_id, idx) => Some(
                self.match_collections
                    .get(mc_id)
                    .and_then(|mc| mc.get(*idx))
                    .map(|m| m.sub_matches.clone())
                    .ok_or_else(|| VbaError::new(91, "Object variable not set")),
            ),
            _ => None,
        }
    }

    pub fn default_value(&self, obj: &ObjRef) -> Option<VResult<Variant>> {
        match obj {
            ObjRef::RegExpMatch(mc_id, idx) => Some(
                self.match_collections
                    .get(mc_id)
                    .and_then(|mc| mc.get(*idx))
                    .map(|m| Variant::Str(m.value.clone()))
                    .ok_or_else(|| VbaError::new(91, "Object variable not set")),
            ),
            _ => None,
        }
    }
}
