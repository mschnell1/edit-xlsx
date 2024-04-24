use std::cmp;
use std::io::BufWriter;
use std::io::Error;
use std::io::Write;
use std::path::Path;

use std::fs::File;

use edit_xlsx::Read;
use edit_xlsx::WorkSheet;
use edit_xlsx::WorkSheetCol;
use edit_xlsx::{FormatColor, Workbook, WorkbookResult};
use std::ops::RangeInclusive;


use std::io::{ErrorKind};

pub fn error_text(text: &str) -> Error {
    Error::new(ErrorKind::Other, text)
}

pub fn tab_header(widths: &Vec<f64>) -> String {
    let mut out = "[cols=\"".to_owned();

    for w in widths.iter().enumerate() {
        let w100 = w.1 * 100.0;
        let w100 = w100.round();
        let w100 = w100 as u32;
        out += &format!("{}", w100);
        if w.0 < widths.len() - 1 {
            out += ", ";
        };
    }
    out + "\"]\r"
}

const CARGO_PKG_NAME: &str = env!("CARGO_PKG_NAME");
fn main() -> std::io::Result<()> {
    let (in_file_name, out_file_name) = (
        // Path::new("./tests/xlsx/yearly-calendar.xlsx"),
        // Path::new("./tests/xlsx/business-budget.xlsx"),
        Path::new("./tests/xlsx/accounting.xlsx"),
        Path::new("./examples/_test.adoc"),
        
    );

    println!(
        "{} {} -> {}",
        CARGO_PKG_NAME,
        in_file_name.display(),
        out_file_name.display()
    );

    xlsx_convert(&in_file_name, &out_file_name)?;
    Ok(())
}



macro_rules! print_debug {
    ($f: expr, $($a: expr),*) => {
//        println!($f, $($a),*);
    };
}

#[derive(Default)]
pub struct Xlsx2AdocTestResults {
    // Todo
    v1: u8,
    pub v2: u8,
}

pub(crate) fn to_col(col: &str) -> u32 {
    let mut col = col.as_bytes();
    let mut num = 0;
    while !col.is_empty() {
        if col[0] > 64 && col[0] < 91 {
            num *= 26;
            num += (col[0] - 64) as u32;
        }
        col = &col[1..];
    }
    num
}

fn decode_col_range(column_name: &str, length: usize) -> RangeInclusive<usize> {
    let mut nn = column_name.split(':');
    let nl = nn.next();
    let nl = nl.unwrap();
    let cl = to_col(nl) as usize;
    let nh = nn.next();
    let nh = nh.unwrap();
    let ch = cmp::min(to_col(nh) as usize, length);
    cl - 1..=ch - 1
}

#[derive(Default)]
struct Cell {
    text: String,
    text_color: FormatColor,
    bg_color: FormatColor,
    bg_bg_color: FormatColor,
}

struct Rgb {
    r: u8,
    g: u8,
    b: u8,
}
struct ColorName<'a> {
    name: &'a str,
    rgb: Rgb,
}

struct ColorNames<'a> {
    color_names: Vec<ColorName<'a>>,
}

impl<'a> ColorNames<'a> {
    fn push<'b>(color_names: &mut Vec<ColorName<'b>>, name: &'b str, r: u8, g: u8, b: u8) {
        let rgb = Rgb { r, g, b };
        let color_name = ColorName { name, rgb };
        color_names.push(color_name);
    }

    fn init() -> Self {
        /*
           con:folder[role=Aqua] Aqua          000 255 255
           icon:folder[role=Black] Black       000 000 000
           icon:folder[role=Blue] Blue         000 000 255
           icon:folder[role=Fuchsia] Fuchsia   255 000 255
           icon:folder[role=Gray] Gray         128 128 128
           icon:folder[role=Green] Green       000 255 000
           icon:folder[role=Lime] Lime         050 205 050
           icon:folder[role=Maroon] Maroon     128 000 000
           icon:folder[role=Navy] Navy         000 000 128
           icon:folder[role=Olive] Olive       186 184 108
           icon:folder[role=Purple] Purple     128 000 128
           icon:folder[role=Red] Red           255 000 000
           icon:folder[role=Silver] Silver     192 192 192
           icon:folder[role=Teal] Teal         000 128 128
           icon:folder[role=White] White       255 255 255
           icon:folder[role=Yellow] Yellow     255 255 000
        */

        let mut color_names = Vec::<ColorName>::new();

        macro_rules! push_color {
            ($n:expr, $r:expr, $g:expr, $b:expr) => {
                Self::push(&mut color_names, $n, $r, $g, $b);
            };
        }

        push_color!("aqua", 0, 255, 255);
        push_color!("black", 0, 0, 0);
        push_color!("blue", 0, 0, 255);
        push_color!("fuchsia", 255, 0, 255);
        push_color!("gray", 128, 128, 128);
        push_color!("green", 0, 255, 0);
        push_color!("lime", 50, 205, 50);
        push_color!("maroon", 128, 0, 0);
        push_color!("navy", 0, 0, 128);
        push_color!("olive", 186, 184, 108);
        push_color!("purple", 128, 0, 128);
        push_color!("red", 255, 0, 0);
        push_color!("silver", 192, 192, 12);
        push_color!("teal", 0, 128, 128);
        push_color!("white", 255, 255, 255);
        push_color!("yellow", 255, 255, 0);
        Self { color_names }
    }

    fn decode_rgb(&self, r: u8, g: u8, b: u8) -> &str {
        let mut min = 0xFFFF;
        let mut name_min = "";
        for n in &self.color_names {
            let d = r as i32;
            let d = d - (n.rgb.r as i32);
            let d = d * d;
            let mut dd = d;
            let d = g as i32;
            let d = d - (n.rgb.g as i32);
            let d = d * d;
            dd += d;
            let d = b as i32;
            let d = d - (n.rgb.b as i32);
            let d = d * d;
            dd += d;
            if dd < min {
                min = dd;
                name_min = n.name;
            }
        }
        print_debug!("-------------------- {} ({})", name_min, min);
        name_min
    }
}

// const COLOR_NAMES: Vec<ColorName> = define_colors();

struct XlsxConvertor<'a> {
    sheet: &'a WorkSheet,
    writer: BufWriter<&'a mut File>,
    widths: Vec<f64>,
    color_names: &'a ColorNames<'a>,
}

impl<'a> XlsxConvertor<'a> {
    fn new(
        sheet: &'a WorkSheet,
        writer: BufWriter<&'a mut File>,
        color_names: &'a ColorNames<'a>,
    ) -> Self {
        XlsxConvertor {
            sheet,
            writer,
            widths: Vec::<f64>::new(),
            color_names,
        }
    }

    fn find_col_width(&mut self) -> Result<(), Error> {
        let mut widths = Vec::<f64>::new();

        // width 0 results in have the text fit in the field if possible"
        let default_col_width = self.sheet.get_default_column().unwrap_or(0.0);

        print_debug!("Default Col Width: {}", default_col_width);

        for _ in 0..self.sheet.max_column() {
            widths.push(default_col_width);
        }

        let formatted_col_result = self.sheet.get_columns_with_format((1, 1, 1, 16384));
        let formatted_col = match formatted_col_result {
            Ok(f) => f,
            Err(e) => return Err(error_text(&format!("{:?}", e))),
        };

        for w in formatted_col.iter() {
            let column_name = w.0;
            let a = w.1;
            let columns_specs = a.0;
            let column_width = columns_specs.width;
            if let Some(width) = column_width {
                let col_range = decode_col_range(column_name, widths.len());
                for c in col_range {
                    widths[c] = width;
                }
            };
        }
        self.widths = widths;
        Ok(())
    }

    fn write_tab_start(&mut self) -> Result<(), Error> {
        let bounds = "|===\r";
        let line = tab_header(&self.widths);
        self.writer.write_all(line.as_bytes())?;
        self.writer.write_all(bounds.as_bytes())?;
        Ok(())
    }

    fn write_tab_end(&mut self) -> Result<(), Error> {
        let bounds = "|===\r";
        self.writer.write_all(bounds.as_bytes())?;
        Ok(())
    }

    fn write_col_end(&mut self) -> Result<(), Error> {
        self.writer.write_all("\r".as_bytes())?;
        Ok(())
    }

    fn write_row_delimiter(&mut self) -> Result<(), Error> {
        self.writer.write_all("|".as_bytes())?;
        Ok(())
    }

    fn write_cell(&mut self, cell: &Cell) -> Result<(), Error> {
        /*
               let text = match cell.text_color {
                   FormatColor::Default => cell.text.clone(),
                   FormatColor::Index(_i) => cell.text.clone(),
                   FormatColor::RGB(r, g, b) => {
                       let trgb = format!("[{}]#", self.color_names.decode_rgb(r, g, b));
                       trgb + &cell.text + "#"
                   }
                   FormatColor::Theme(_i, _f) => cell.text.clone(),
               };
        */
        let text_color_str = match cell.text_color {
            FormatColor::Default => None,
            FormatColor::Index(_i) => None,
            FormatColor::RGB(r, g, b) => Some(self.color_names.decode_rgb(r, g, b)),
            FormatColor::Theme(_i, _f) => None,
        };
        let bg_color_str = match cell.bg_color {
            FormatColor::Default => None,
            FormatColor::Index(_i) => None,
            FormatColor::RGB(r, g, b) => Some(self.color_names.decode_rgb(r, g, b)),
            FormatColor::Theme(_i, _f) => None,
        };
        let text = match (text_color_str, bg_color_str) {
            (None, None) => cell.text.clone(),
            (Some(tc), None) => format!("[{}]#{}#", tc, cell.text),
            (None, Some(bc)) => format!("[{}-background]#{}#", bc, cell.text),
            (Some(tc), Some(bc)) => format!("[{} {}-background]#{}#", tc, bc, cell.text),
        };

        self.writer.write_all(text.as_bytes())?;
        Ok(())
    }

    fn get_cell(&mut self, row: u32, col: u32) -> Cell {
        let cell_content = self.sheet.read_cell((row + 1, col + 1)).unwrap_or_default();
        let format = cell_content.format;
        let mut cell = Cell::default();
        if format.is_some() {
            let format = format.unwrap();
            cell.text_color = *format.get_color();
            let ff = format.get_background().clone();
            cell.bg_color = ff.fg_color;
            cell.bg_bg_color = ff.bg_color;
        }
        /**/
        let cell_format_string = format!(
            "Text-Color = {:?}        bg = {:?}        bg_bg = {:?}",
            cell.text_color, cell.bg_color, cell.bg_bg_color
        );
        /**/
        let cell_text = cell_content.text;
        cell.text = match cell_text {
            Some(t) => t,
            None => "".to_owned(),
        };
        /**/
        print_debug!(
            "{} ({}) -> {}     Format: {}",
            col,
            self.widths[(col) as usize],
            cell.text,
            cell_format_string
        );
        cell
    }

    fn write_cells(&mut self) -> Result<(), Error> {
        for row in 0..self.sheet.max_row() {
            /*
                   print_debug!("Row {} ({})", row, hights[row as usize]);
            */
            for col in 0..self.sheet.max_column() {
                if col < self.sheet.max_column() {
                    self.write_row_delimiter()?;
                }
                let cell = self.get_cell(row, col);
                self.write_cell(&cell)?;
            }
            self.write_col_end()?;
        }
        Ok(())
    }
}

pub fn xlsx_convert(
    in_file_name: &Path,
    out_file_name: &Path,
) -> Result<Xlsx2AdocTestResults, Error> {
    let workbook = Workbook::from_path(in_file_name);
    let mut workbook = workbook.unwrap();
    workbook.finish();

    let reading_sheet = workbook.get_worksheet(1);
    let sheet = match reading_sheet {
        Ok(s) => s,
        Err(e) => return Err(error_text(&format!("{:?}", e))),
    };
    /*
       let default_row_hight = sheet.get_default_row();
       let mut hights = Vec::<f64>::new();
       for _ in 0..sheet.max_row() {
           hights.push(default_row_hight);
       }
    */
    let mut output_file = File::create(out_file_name)?; // overwrites existing file
    let writer = BufWriter::new(&mut output_file);

    let color_names = ColorNames::init();

    let mut xlsx_converter = XlsxConvertor::new(sheet, writer, &color_names);
    xlsx_converter.find_col_width()?;
    xlsx_converter.write_tab_start()?;
    xlsx_converter.write_cells()?;
    xlsx_converter.write_tab_end()?;

    let xlsx_2_adoc_test_results = Xlsx2AdocTestResults { v1: 0, v2: 0 };
    Ok(xlsx_2_adoc_test_results)
}
