import os
from typing import Optional, Dict, NamedTuple

import openpyxl as pyxl
import openpyxl.writer.excel
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet


def merge_spreadsheets(main_file_path: str, new_file_path: str, column_key: str, merged_file_name: str):
    if os.path.realpath(new_file_path) == os.path.realpath(main_file_path):
        raise Exception("Spreadsheet file paths cannot be identical.")

    main_wb: Workbook = pyxl.load_workbook(main_file_path)
    main_wb.active = 0
    main_sheet: Worksheet = main_wb.active

    new_wb: Workbook = pyxl.load_workbook(new_file_path, read_only=True)
    new_wb.active = 0
    new_sheet: Worksheet = new_wb.active

    # Locate column label positions
    label_indices = _locate_labels(main_sheet, new_sheet)
    if column_key not in label_indices:
        raise Exception(f"The key '{column_key}' is not a column label in either of the spreadsheets.")
    if label_indices[column_key].main_label_index is None:
        raise Exception(f"The key '{column_key}' is not a column label in `{main_file_path}`.")
    if label_indices[column_key].new_label_index is None:
        raise Exception(f"The key '{column_key}' is not a column label in `{new_file_path}`.")
    opp_id_col_indices = label_indices[column_key]

    # Find all "Opportunity ID's" in the main sheet
    for new_opp_row in new_sheet.iter_rows(min_row=2, max_row=new_sheet.max_row):
        if __debug__:
            print(f"CURRENT ROW: {new_opp_row[opp_id_col_indices.new_label_index].value}")

        new_opp_id = str(new_opp_row[opp_id_col_indices.new_label_index].value)
        main_opp_row = _locate_opp_row(main_sheet, opp_id_col_indices.main_label_index, new_opp_id)

        # If a row wasn't found for this opportunity in the main sheet, it must be new and inserted into the sheet
        if main_opp_row is None:
            main_sheet.insert_rows(2)
            main_opp_row = next(main_sheet.iter_rows(min_row=2, max_row=2, max_col=len(label_indices)))

        # Update all previous opportunities in the main sheet with new sheet data
        _update_opp_row(main_opp_row, new_opp_row, label_indices)

    # If the merged file name is blank, assume that the main file will be used instead
    if merged_file_name is None or len(merged_file_name) == 0:
        merged_file_path = main_file_path
    else:
        merged_file_dir = os.path.dirname(main_file_path)
        merged_file_ext = os.path.splitext(main_file_path)[1]
        merged_file_path = os.path.join(f"{merged_file_dir}", f"{merged_file_name}{merged_file_ext}")

    # Creates a copy
    openpyxl.writer.excel.save_workbook(main_wb, merged_file_path)

    # Close workbook files
    main_wb.close()
    new_wb.close()


class _LabelIndices(NamedTuple):
    main_label_index: int
    new_label_index: Optional[int]


def _locate_labels(main_sheet: Worksheet, new_sheet: Worksheet):
    """Map column label indexes used in new sheet to the main sheet indexes"""
    label_indices = {}

    # Iterate over all the column labels in the main sheet
    for (main_label_cell,) in main_sheet.iter_cols(max_row=1):
        # if the main column label is empty, skip it
        if main_label_cell.value is None:
            continue

        # Find new sheet position of current column label being examined in the main sheet
        for new_col in range(1, new_sheet.max_column + 1):
            new_label_cell = new_sheet.cell(1, new_col)

            # if the new column label is empty, skip it
            if new_label_cell.value is None:
                continue

            # Keep track of where the label exists in each
            if main_label_cell.value.lower() == new_label_cell.value.lower():
                label_indices[main_label_cell.value] = _LabelIndices(main_label_cell.column - 1,
                                                                     new_label_cell.column - 1)
                break

            new_col += 1

        # Column label in main does not exist in new, and should be ignored later on
        if main_label_cell.value not in label_indices:
            label_indices[main_label_cell.value] = _LabelIndices(main_label_cell.column, None)

    return label_indices


def _locate_opp_row(sheet: pyxl.workbook.workbook.Worksheet,
                    opp_id_col_index: int, opp_id: str) -> Optional[int]:
    """Using the given opportunity id, output the row data it is located on within the opportunity id
    column of the given worksheet"""
    for curr_opp_row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if str(curr_opp_row[opp_id_col_index].value) == opp_id:
            return curr_opp_row

    # The given opportunity id wasn't found
    return None


def _update_opp_row(main_row, new_row, label_indices: Dict[str, _LabelIndices]):
    for indexes in label_indices.values():
        # If label does not exist in the new sheet, just ignore this label
        if indexes.new_label_index is None:
            continue

        main_row[indexes.main_label_index].value = new_row[indexes.new_label_index].value
