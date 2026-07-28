use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct Position {
    pub left: isize,
    pub top: isize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct Rect {
    pub top: isize,
    pub left: isize,
    pub width: usize,
    pub height: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct TextPosition {
    pub top: isize,
    pub left: isize,
    pub char_offset: usize,
}

impl TextPosition {
    pub fn to_pos(&self) -> Position {
        Position {
            left: self.left,
            top: self.top,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct UnitValues {
    pub char_dims: (f32, f32),
    pub tile_dims: (f32, f32),
}
