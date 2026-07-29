use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ResultData {
    None,
    Boolean(bool),
    Integer(i64),
    Float(f64),
    String(String),
    List(Vec<ResultData>),
    Dict(Vec<(ResultData, ResultData)>),
    Plot {
        points: Vec<(f32, f32)>,
        color: [f32; 4],
        radius: f32,
        is_line: bool,
        title: Option<String>,
        xlabel: Option<String>,
        ylabel: Option<String>,
    },
    Error(String),
}

impl ResultData {
    pub fn to_string(&self) -> String {
        match self {
            ResultData::None => "".to_string(),
            ResultData::Boolean(b) => {
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            ResultData::Integer(i) => format!("{}", i),
            ResultData::Float(f) => format!("{}", f),
            ResultData::String(s) => s.clone(),
            ResultData::List(l) => {
                let items: Vec<String> = l.iter().map(|i| i.to_string()).collect();
                format!("[{}]", items.join(", "))
            }
            ResultData::Dict(d) => {
                let items: Vec<String> = d
                    .iter()
                    .map(|(k, v)| format!("{}: {}", k.to_string(), v.to_string()))
                    .collect();
                format!("{{ {} }}", items.join(", "))
            }
            ResultData::Plot { .. } => "[Plot]".to_string(),
            ResultData::Error(e) => format!("Error: {}", e),
        }
    }

    pub fn plot_cell_dims(&self) -> Option<(usize, usize)> {
        if let ResultData::Plot { points, .. } = self {
            if points.is_empty() {
                return Some((16, 16));
            }
            let mut x_min = f32::INFINITY;
            let mut x_max = f32::NEG_INFINITY;
            let mut y_min = f32::INFINITY;
            let mut y_max = f32::NEG_INFINITY;
            for &(x, y) in points {
                if x < x_min {
                    x_min = x;
                }
                if x > x_max {
                    x_max = x;
                }
                if y < y_min {
                    y_min = y;
                }
                if y > y_max {
                    y_max = y;
                }
            }
            if !x_min.is_finite() || !x_max.is_finite() || !y_min.is_finite() || !y_max.is_finite()
            {
                return Some((16, 16));
            }
            let dx = x_max - x_min;
            let dy = y_max - y_min;
            if dx == 0.0 && dy == 0.0 {
                return Some((16, 16));
            }

            // Proportional scaling: the larger of the two dimensions should be 15
            let w;
            let h;
            if dx >= dy {
                w = 15;
                if dx > 0.0 {
                    h = ((15.0 * dy / dx).round() as usize).max(1);
                } else {
                    h = 1;
                }
            } else {
                h = 15;
                if dy > 0.0 {
                    w = ((15.0 * dx / dy).round() as usize).max(1);
                } else {
                    w = 1;
                }
            }
            // Add a column to the left and a row to the bottom for axis labels
            Some((w + 1, h + 1))
        } else {
            None
        }
    }
}
