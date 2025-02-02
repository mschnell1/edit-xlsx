//! Some traits for managing deserialized Cells
//! - Add, modify and delete Cells.
use serde::{Deserialize, Serialize};
use crate::api::cell::values::{CellDisplay, CellValue};
use crate::api::worksheet::row::Row as ApiRow;
use crate::api::cell::Cell as ApiCell;
use crate::result::RowResult;
use crate::xml::worksheet::sheet_data::Cell;

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
pub(crate) struct Row {
    // Cells are ordered
    #[serde(rename = "c", default, skip_serializing_if = "Vec::is_empty")]
    pub(crate) cells: Vec<Cell>,
    #[serde(rename = "@r")]
    pub(crate) row: u32,
    #[serde(rename = "@spans", skip_serializing_if = "Option::is_none")]
    pub(crate) spans: Option<String>,
    #[serde(rename = "@s", skip_serializing_if = "Option::is_none")]
    pub(crate) style: Option<u32>,
    #[serde(rename = "@customFormat", skip_serializing_if = "Option::is_none")]
    pub(crate) custom_format: Option<u8>,
    #[serde(rename = "@ht", skip_serializing_if = "Option::is_none")]
    pub(crate) height: Option<f64>,
    #[serde(rename = "@customHeight", skip_serializing_if = "Option::is_none")]
    pub(crate) custom_height: Option<u8>,
    #[serde(rename = "@outlineLevel", skip_serializing_if = "Option::is_none")]
    pub(crate) outline_level: Option<u32>,
    #[serde(rename = "@collapsed", skip_serializing_if = "Option::is_none")]
    pub(crate) collapsed: Option<u8>,
    #[serde(rename = "@hidden", skip_serializing_if = "Option::is_none")]
    pub(crate) hidden: Option<u8>,
    #[serde(rename(serialize = "@x14ac:dyDescent", deserialize = "@dyDescent"), skip_serializing_if = "Option::is_none")]
    pub(crate) x14ac_dy_descent: Option<f64>,
}

impl Row {
    pub(crate) fn to_api_row(&self) -> ApiRow {
        let mut row = ApiRow::default();
        row.height = self.height;
        row.style = self.style;
        row.hidden = self.hidden;
        row.outline_level = self.outline_level;
        row.collapsed = self.collapsed;
        row
    }
}

impl Row {
    pub(crate) fn new(row: u32) -> Row {
        Row {
            cells: vec![],
            row,
            height: None,
            spans: None,
            style: None,
            custom_format: None,
            x14ac_dy_descent: None,
            custom_height: None,
            outline_level: None,
            collapsed: None,
            hidden: None,
        }
    }

    pub(crate) fn max_col(&self) -> u32 {
        match self.cells.last() {
            Some(cell) => cell.loc.col,
            None => 0
        }
    }

    pub(crate) fn add_by_api_cell<T: CellDisplay + CellValue>(&mut self, col: u32, api_cell: &ApiCell<T>) -> RowResult<()> {
        let cell = self.get_or_new_cell(col);
        cell.update_by_api_cell(api_cell)?;
        Ok(())
    }

    pub(crate) fn add_display_cell<T: CellDisplay + CellValue>(&mut self, col: u32, text: &T, style: Option<u32>) {
        // 判断新增cell位置是否已经存在别的cell
        let cell = self.get_or_new_cell(col);
        cell.update_by_display(text, style);
    }

    pub(crate) fn get_display_cell(&self, col: u32) -> Option<&String> {
        let cell = self.get_cell(col);
        if let Some(cell) = cell {
            // return if let Some(formula) = &cell.formula {
            //     Some(formula.get_formula())
            // } else {
            cell.text.as_ref()
            // }
        } else {
            None
        }
    }

    // pub(crate) fn add_formula_cell(&mut self, col: u32, formula: &str, formula_type: FormulaType, style: Option<u32>) {
    //     let cell = self.get_or_new_cell(col);
    //     cell.update_by_formula(formula, formula_type, style);
    // }
}

pub(crate) trait _OrderCell {
    fn get_position_by_col(&self, col: u32) -> usize;
    fn new_cell(&mut self, col: u32) -> &mut Cell;
    fn get_cell(&self, col: u32) -> Option<&Cell>;
    fn get_cell_mut(&mut self, col: u32) -> Option<&mut Cell>;
    fn get_or_new_cell(&mut self, col: u32) -> &mut Cell;
}

impl _OrderCell for Row {
    fn get_position_by_col(&self, col: u32) -> usize {
        let mut l = 0;
        let mut r = self.cells.len();
        while r - l > 0 {
            let mid = (l + r) / 2;
            if col == self.cells[mid].loc.col {
                return mid;
            } else if col < self.cells[mid].loc.col {
                r = mid;
            } else {
                l = mid + 1;
            }
        }
        r
    }

    fn new_cell(&mut self, col: u32) -> &mut Cell {
        let r = self.get_position_by_col(col);
        self.cells.insert(r, Cell::new((self.row, col)));
        return &mut self.cells[r];
    }

    fn get_cell(&self, col: u32) -> Option<&Cell> {
        let r = self.get_position_by_col(col);
        if r >= self.cells.len() {return None}
        return if col == self.cells[r].loc.col { Some(&self.cells[r]) } else { None }
    }

    fn get_cell_mut(&mut self, col: u32) -> Option<&mut Cell> {
        let r = self.get_position_by_col(col);
        if r >= self.cells.len() {return None}
        return if col == self.cells[r].loc.col { Some(&mut self.cells[r]) } else { None }
    }

    fn get_or_new_cell(&mut self, col: u32) -> &mut Cell {
        let r = self.get_position_by_col(col);
        return if r < self.cells.len() && self.cells[r].loc.col == col {
            &mut self.cells[r]
        } else {
            self.cells.insert(r, Cell::new((self.row, col)));
            &mut self.cells[r]
        }
    }
}
