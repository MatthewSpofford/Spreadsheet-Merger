import os
from typing import Optional, Dict

import openpyxl as pyxl


_OPPORTUNITY_LABEL = "Opportunity Id"


def merge_spreadsheets(main_file_path, new_file_path):
    main_wb: pyxl.workbook.workbook.Workbook = pyxl.load_workbook(main_file_path)
    main_sheet: pyxl.workbook.workbook.Worksheet = \
        main_wb["Manager_Opportunity_Dashboard"]
    new_sheet: pyxl.workbook.workbook.Worksheet = \
        pyxl.load_workbook(new_file_path, read_only=True)["Manager_Opportunity_Dashboard"]

    # Locate column label positions
    label_positions = _locate_labels(main_sheet, new_sheet)
    opp_id_col = label_positions[_OPPORTUNITY_LABEL].new_col
    opp_id_col_index = opp_id_col - 1

    # Find all "Opportunity ID's" in the main sheet
    for (new_opp_row,) in new_sheet.iter_rows(min_col=opp_id_col, max_col=opp_id_col, min_row=2):
        new_opp_id = str(new_opp_row[opp_id_col_index].value)
        main_opp_row = _locate_opp_row(main_sheet, opp_id_col, opp_id_col_index, new_opp_id)

        # If a row wasn't found for this opportunity in the main sheet, it must be new and inserted into the sheet
        if not main_opp_row:
            main_sheet.insert_rows(2)
            main_opp_row = main_sheet.iter_rows(min_row=2, max_row=2)

        # Update all previous opportunities in the main sheet with new sheet data
        _update_opp_row(main_opp_row, new_opp_row, label_positions)

    # TODO REMOVE THIS FOR FINAL VERSION
    # Creates copy
    new_file_path = os.path.splitext(main_file_path)
    new_file_path = os.path.join(new_file_path[0] + " (Copy)", new_file_path[1])
    main_wb.save(new_file_path)


class _LabelColumns:
    def __init__(self, main_col: int, new_col: Optional[int]):
        self.main_col = main_col
        self.new_col = new_col


def _locate_labels(main_sheet: pyxl.workbook.workbook.Worksheet, new_sheet: pyxl.workbook.workbook.Worksheet):
    """Map column labels used in new sheet to the main sheet positions"""
    label_positions = {}

    for (main_label_cell,) in main_sheet.iter_cols(min_col=1, min_row=1, max_row=1):
        new_col = 1

        # Find new sheet position of current column label being examined in the main sheet
        while new_sheet.cell(1, new_col).value is not None:
            new_label_cell = new_sheet.cell(1, new_col)

            # Keep track of where the label exists in each
            if main_label_cell.value.lower() == new_label_cell.value.lower():
                label_positions[main_label_cell.value] = _LabelColumns(main_label_cell.column, new_label_cell.column)
                break

            new_col += 1

        # Column label in main does not exist in new, and should be ignored later on
        if main_label_cell.value not in label_positions:
            label_positions[main_label_cell.value] = _LabelColumns(main_label_cell.column, None)

    return label_positions


def _locate_opp_row(sheet: pyxl.workbook.workbook.Worksheet,
                    opp_id_col: int, opp_id_col_index: int, opp_id: str)-> Optional[int]:
    """Using the given opportunity id, output the row data it is located on within the opportunity id
    column of the given worksheet"""
    for curr_opp_row in sheet.iter_rows(min_col=opp_id_col, max_col=opp_id_col, min_row=2):
        if curr_opp_row[opp_id_col_index].value == opp_id:
            return curr_opp_row

    # The given opportunity id wasn't found
    return None


def _update_opp_row(main_row, new_row, label_positions: Dict[str, _LabelColumns]):
    pass
