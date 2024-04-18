//!
//! This module contains the [`FormatFont`] struct, which used to set the fill of a format.
//! [`FormatFont`] is used to handle font-related fields, mainly applied to [`Cell`].
//! # Examples
//!
//! Use [`FormatFont`] to set text's font name, size, and color
//! ```
//! use edit_xlsx::{Format, FormatColor, FormatFill, Workbook, WorkSheetCol, Write};
//! let mut workbook = Workbook::from_path("./examples/xlsx/accounting.xlsx").unwrap();
//! let worksheet = workbook.get_worksheet_mut_by_name("worksheet").unwrap();
//! let mut format = Format::default();
//! format.font.name = "Wide Latin".to_string();
//! format.font.size = 30.0;
//! format.font.color = FormatColor::Index(4);
//! worksheet.write_with_format("A1", "It's so fun", &format).unwrap();
//! workbook.save_as("./examples/font_write_with_font.xlsx").unwrap();
//! ```

use crate::{Format, FormatColor};
use crate::xml::common::{Element, FromFormat};
use crate::xml::style::color::Color;
use crate::xml::style::font::{Bold, Font, Italic, Underline};
use crate::Cell;

///
/// [`FormatFont`] is used to handle font-related fields, mainly applied to [`Cell`].
/// # Fields
/// | field     | type        | meaning                                |
/// | --------- | ----------- | -------------------------------------- |
/// | `bold`      | bool        | Whether the font is bold.              |
/// | `italic`    | bool        | Whether the font is italic.            |
/// | `underline` | bool        | Whether the font is underlined.        |
/// | `size`      | f64         | Font size in point.                    |
/// | `color`     | [`FormatColor`] | Font color.                            |
/// | `name`      | [`String`]      | Font name, same as displayed in Excel. |
///
#[derive(Clone, Debug, PartialEq)]
pub struct FormatFont {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub size: f64,
    pub color: FormatColor,
    pub name: String,
}

impl Default for FormatFont {
    ///
    /// Method [`FormatFont::default()`] creates a default [`FormatFont`] in Excel,
    /// with font name *Calibri* and font size 11.0 in point
    ///
    /// When using [`Format::default()`], its `font` field will be filled by this method.
    fn default() -> Self {
        FormatFont {
            bold: false,
            italic: false,
            underline: false,
            size: 11.0,
            color: Default::default(),
            name: "Calibri".to_string(),
        }
    }
}

impl FromFormat<FormatFont> for Font {
    fn set_attrs_by_format(&mut self, format: &FormatFont) {
        self.color = Some(Color::from_format(&format.color));
        self.name = Some(Element::from_val(format.name.to_string()));
        self.sz = Some(Element::from_val(format.size));
        self.bold = if format.bold { Some(Bold::default()) } else { None };
        self.underline = if format.underline { Some(Underline::default()) } else { None };
        self.italic = if format.italic { Some(Italic::default()) } else { None };
    }

    fn set_format(&self, format: &mut FormatFont) {
        format.bold = self.bold.is_some();
        format.italic = self.italic.is_some();
        format.underline = self.underline.is_some();
        if let Some(size) = &self.sz {
            format.size = size.get_format();
        }
        if let Some(name) = &self.name {
            format.name = name.val.to_string();
        }
        format.color = self.color.as_ref().get_format();
    }
}
