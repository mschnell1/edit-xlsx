use serde::{Deserialize, Serialize};
use crate::api::format::FormatAlign;
use crate::FormatAlignType;
use crate::xml::common::FromFormat;

#[derive(Debug, Deserialize, Serialize, PartialEq, Clone, Default)]
pub(crate) struct Alignment {
    #[serde(rename = "@horizontal", skip_serializing_if = "Option::is_none")]
    pub(crate) horizontal: Option<String>,
    #[serde(rename = "@vertical", skip_serializing_if = "Option::is_none")]
    pub(crate) vertical: Option<String>,
    #[serde(rename = "@textRotation", skip_serializing_if = "Option::is_none")]
    text_rotation: Option<u8>,
    #[serde(rename = "@wrapText", skip_serializing_if = "Option::is_none")]
    wrap_text: Option<u8>,
    #[serde(rename = "@indent", skip_serializing_if = "Option::is_none")]
    indent: Option<u8>,
    #[serde(rename = "@justifyLastLine", skip_serializing_if = "Option::is_none")]
    justify_last_line: Option<u8>,
    #[serde(rename = "@shrinkToFit", skip_serializing_if = "Option::is_none")]
    shrink_to_fit: Option<u8>,
    #[serde(rename = "@readingOrder", skip_serializing_if = "Option::is_none")]
    reading_order: Option<u8>,
}

// impl Alignment {
//     pub(crate) fn is_empty(&self) -> bool {
//         self.horizontal.is_none()
//             &&self.vertical.is_none()
//             &&self.text_rotation.is_none()
//             &&self.wrap_text.is_none()
//             &&self.indent.is_none()
//             &&self.justify_last_line.is_none()
//             &&self.shrink_to_fit.is_none()
//             &&self.reading_order.is_none()
//     }
// }

impl FromFormat<FormatAlign> for Alignment {
    fn set_attrs_by_format(&mut self, format: &FormatAlign) {
        if let Some(vertical) = format.vertical {
            self.vertical = Some(String::from(vertical.to_str()));
        }
        if let Some(horizontal) = format.horizontal {
            self.horizontal = Some(String::from(horizontal.to_str()));
        }
        self.reading_order = format.reading_order;
        self.indent = format.indent;
    }

    fn set_format(&self, format: &mut FormatAlign) {
        format.indent = self.indent;
        format.reading_order = self.reading_order;
        format.horizontal = FormatAlignType::from_str(self.horizontal.as_ref(), true);
        format.vertical = FormatAlignType::from_str(self.horizontal.as_ref(), false);
    }
}

// impl Alignment {
//     pub(crate) fn new(horizontal: Option<String>, vertical: Option<String>) -> Alignment {
//         Alignment {
//             horizontal,
//             vertical,
//         }
//     }
// }