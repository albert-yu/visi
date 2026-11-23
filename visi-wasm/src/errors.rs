use js_sys::{Array, Error as JsError, Reflect};
use visi_engine::Error as CoreError;
use wasm_bindgen::JsValue;

pub type WasmResult<T> = Result<T, JsValue>;

pub struct Wrapped(pub CoreError);

impl From<CoreError> for Wrapped {
    fn from(e: CoreError) -> Self {
        Wrapped(e)
    }
}

fn set_prop(target: &JsValue, key: &str, val: JsValue) {
    let _ = Reflect::set(target, &JsValue::from_str(key), &val);
}

impl From<Wrapped> for JsValue {
    fn from(w: Wrapped) -> JsValue {
        let e = w.0;
        let msg = e.to_string();
        let err_name = match &e {
            CoreError::NotFound { .. } => "NotFoundError",
            CoreError::AlreadyExists { .. } => "AlreadyExistsError",
            CoreError::NameTaken { .. } => "NameTakenError",
            CoreError::InvalidName { .. } => "InvalidNameError",
            CoreError::OutOfBounds { .. } => "OutOfBoundsError",
            CoreError::InvalidRange(_) => "InvalidRangeError",
            CoreError::InvalidArgument(_) => "InvalidArgumentError",
            CoreError::Csv(_) => "CsvError",
            CoreError::Xlsx(_) => "XlsxError",
            CoreError::Xlsb(_) => "XlsbError",
            CoreError::Vba(_) => "VbaError",
            CoreError::VbaSyntax { .. } => "VbaSyntaxError",
            CoreError::VbaRuntime { .. } => "VbaRuntimeError",
            CoreError::Eval(_) => "EvaluationError",
            CoreError::EmptyWorkbook => "EmptyWorkbookError",
            CoreError::LastSheetInWorkbook => "LastSheetError",
            CoreError::DocumentModuleExists => "DocumentModuleExistsError",
            CoreError::Io(_) => "IoError",
            _ => "VisiError",
        };

        let js_err = JsError::new(&msg);
        js_err.set_name(err_name);
        let obj: JsValue = js_err.into();
        set_prop(&obj, "code", JsValue::from_str(err_name));

        match &e {
            CoreError::NotFound {
                kind,
                name,
                available,
            } => {
                set_prop(&obj, "kind", JsValue::from_str(kind.as_str()));
                set_prop(&obj, "target", JsValue::from_str(name.as_str()));
                set_prop(&obj, "objectName", JsValue::from_str(name.as_str()));
                let arr = Array::new();
                for item in available {
                    arr.push(&JsValue::from_str(item));
                }
                set_prop(&obj, "available", arr.into());
            }
            CoreError::AlreadyExists { kind, name } | CoreError::NameTaken { kind, name } => {
                set_prop(&obj, "kind", JsValue::from_str(kind.as_str()));
                set_prop(&obj, "target", JsValue::from_str(name.as_str()));
                set_prop(&obj, "objectName", JsValue::from_str(name.as_str()));
            }
            CoreError::InvalidName { kind, name, reason } => {
                set_prop(&obj, "kind", JsValue::from_str(kind.as_str()));
                set_prop(&obj, "target", JsValue::from_str(name.as_str()));
                set_prop(&obj, "objectName", JsValue::from_str(name.as_str()));
                set_prop(&obj, "reason", JsValue::from_str(reason.as_str()));
            }
            CoreError::VbaRuntime { number, .. } => {
                set_prop(&obj, "number", JsValue::from_f64(*number as f64));
            }
            CoreError::VbaSyntax {
                module,
                line,
                column,
                ..
            } => {
                let mod_val = match module {
                    Some(m) => JsValue::from_str(m),
                    None => JsValue::NULL,
                };
                set_prop(&obj, "module", mod_val);
                set_prop(&obj, "line", JsValue::from_f64(*line as f64));
                set_prop(&obj, "column", JsValue::from_f64(*column as f64));
            }
            CoreError::OutOfBounds { what, index, len } => {
                set_prop(&obj, "what", JsValue::from_str(what));
                set_prop(&obj, "index", JsValue::from_f64(*index as f64));
                set_prop(&obj, "length", JsValue::from_f64(*len as f64));
            }
            _ => {}
        }

        obj
    }
}

pub fn invalid_argument(msg: impl std::fmt::Display) -> JsValue {
    Wrapped(CoreError::InvalidArgument(msg.to_string())).into()
}
