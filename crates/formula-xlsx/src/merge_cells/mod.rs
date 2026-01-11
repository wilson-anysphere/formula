mod read;
mod write;

pub use read::{read_merge_cells_from_worksheet_xml, read_merge_cells_from_xlsx, MergeCellsError};
pub use write::{update_worksheet_xml, write_merge_cells_section, write_worksheet_xml};
