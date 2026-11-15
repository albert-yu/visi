use js_sys::{Array, Map, Object, Reflect};
use visi_engine::core::ResultData;
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellError {
    code: String,
}

#[wasm_bindgen]
impl CellError {
    #[wasm_bindgen(constructor)]
    pub fn new(code: String) -> Self {
        Self { code }
    }

    #[wasm_bindgen(getter)]
    pub fn code(&self) -> String {
        self.code.clone()
    }

    #[wasm_bindgen(js_name = toString)]
    pub fn to_string_js(&self) -> String {
        self.code.clone()
    }

    #[wasm_bindgen(js_name = toJSON)]
    pub fn to_json_js(&self) -> JsValue {
        let obj = Object::new();
        let _ = Reflect::set(
            &obj,
            &JsValue::from_str("type"),
            &JsValue::from_str("error"),
        );
        let _ = Reflect::set(
            &obj,
            &JsValue::from_str("code"),
            &JsValue::from_str(&self.code),
        );
        obj.into()
    }
}

pub fn result_to_js(v: &ResultData) -> JsValue {
    match v {
        ResultData::None => JsValue::NULL,
        ResultData::Boolean(b) => JsValue::from_bool(*b),
        ResultData::Integer(i) => JsValue::from_f64(*i as f64),
        ResultData::Float(f) => JsValue::from_f64(*f),
        ResultData::String(s) => JsValue::from_str(s),
        ResultData::List(items) => {
            let arr = Array::new_with_length(items.len() as u32);
            for (idx, item) in items.iter().enumerate() {
                arr.set(idx as u32, result_to_js(item));
            }
            arr.into()
        }
        ResultData::Dict(pairs) => {
            let map = Map::new();
            for (k, val) in pairs {
                map.set(&result_to_js(k), &result_to_js(val));
            }
            map.into()
        }
        ResultData::Error(code) => CellError::new(code.clone()).into(),
    }
}
