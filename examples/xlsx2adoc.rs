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

fn find_col_width(sheet: &WorkSheet) -> Result<Vec<f64>, Error> {
    let mut widths = Vec::<f64>::new();
    let default_col_width = sheet.get_default_column().unwrap_or(1.0);

    for _ in 0..sheet.max_column() {
        widths.push(default_col_width);
    }

    let formatted_col_result = sheet.get_columns_with_format((1, 1, 1, 16384));
    let formatted_col = match formatted_col_result {
        Ok(f) => f,
        Err(e) => return Err(error_text(&format!("{:?}", e))),
    };

    for w in formatted_col.iter() {
        let column_name = w.0;
        let a = w.1;
        let columns_specs = a.0;
        let column_width = columns_specs.width;
        match column_width {
            Some(width) => {
                let col_range = decode_col_range(column_name, widths.len());
                for c in col_range {
                    widths[c] = width;
                }
            }
            None => {},
        };
}
    Ok(widths)
}

#[derive(Default)]
struct Cell {
    text: String,
    text_color: FormatColor,
    bg_color: FormatColor,
    bg_bg_color: FormatColor,
}

fn get_cell(row: u32, col: u32, sheet: &WorkSheet, widths: &Vec<f64>) -> Cell {
    let cell_content = sheet.read_cell((row + 1, col + 1)).unwrap_or_default();
    let format = cell_content.format;
    let mut cell = Cell::default();
    if format.is_some() {
        let format = format.unwrap();
        cell.text_color = *format.get_color();
        let ff = format.get_background().clone();
        cell.bg_color = ff.fg_color;
        cell.bg_bg_color = ff.bg_color;
    }

    let cell_text = cell_content.text;
    cell.text = match cell_text {
        Some(t) => t,
        None => "".to_owned(),
    };
    cell
}

fn write_tab_start(writer: &mut BufWriter<&mut File>, widths: &Vec<f64>) -> Result<(), Error> {
    let bounds = "|===\r";
    let line = tab_header(&widths);
    writer.write_all(line.as_bytes())?;
    writer.write_all(bounds.as_bytes())?;
    Ok(())
}

fn write_tab_end(writer: &mut BufWriter<&mut File>) -> Result<(), Error> {
    let bounds = "|===\r";
    writer.write_all(bounds.as_bytes())?;
    Ok(())
}

fn write_row_delimiter(writer: &mut BufWriter<&mut File>) -> Result<(), Error> {
    writer.write_all("|".as_bytes())?;
    Ok(())
}


fn write_col_end(writer: &mut BufWriter<&mut File>) -> Result<(), Error> {
    writer.write_all("\r".as_bytes())?;
    Ok(())
}

fn write_cell(cell: &Cell, writer: &mut BufWriter<&mut File>) -> Result<(), Error> {
    writer.write_all(cell.text.as_bytes())?;
    Ok(())
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
    let widths = find_col_width(sheet)?;

    let mut output_file = File::create(out_file_name)?; // overwrites existing file
    let mut writer = BufWriter::new(&mut output_file);

    write_tab_start(&mut writer, &widths)?;

    for row in 0..sheet.max_row() {
        /*
               println!("Row {} ({})", row, hights[row as usize]);
        */
        for col in 0..sheet.max_column() {
            if col < sheet.max_column() {
                write_row_delimiter(&mut writer)?;
            }

            let cell = get_cell(row, col, sheet, &widths);
            write_cell(&cell, &mut writer)?;
        }

        write_col_end(&mut writer)?;
    }
    write_tab_end(&mut writer)?;

    let xlsx_2_adoc_test_results = Xlsx2AdocTestResults { v1: 0, v2: 0 };
    Ok(xlsx_2_adoc_test_results)
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
