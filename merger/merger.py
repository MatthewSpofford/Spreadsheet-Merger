import os
from typing import Optional

import openpyxl as pyxl


_OPPORTUNITY_LABEL = "Opportunity Value"


def merge_spreadsheets(main_file_path, new_file_path):
    main_wb = pyxl.load_workbook(main_file_path)
    main_sheet = main_wb["Manager_Opportunity_Dashboard"]
    new_sheet = pyxl.load_workbook(new_file_path, read_only=True)["Manager_Opportunity_Dashboard"]

    label_positions = _locate_labels(main_sheet, new_sheet)

    # Update all previous opportunities in the main sheet with new sheet data
    # - only copy values and keep formatting
    row = 2
    while main_sheet.cell(row, label_positions[_OPPORTUNITY_LABEL].main_cell).value is not None:
        pass

    # Find all "Opportunity ID's" in the main sheet

    # Insert new opportunities from at the top of the main sheet

    # TODO REMOVE THIS FOR FINAL VERSION
    # Creates copy
    path_split = os.path.splitext(main_file_path)
    main_wb.save(os.path.join(path_split[0] + " (Copy)", path_split[1]))


class _LabelColumns:
    def __init__(self, main_cell: int, new_cell: Optional[int]):
        self.main_cell = main_cell
        self.new_cell = new_cell


def _locate_labels(main_sheet: pyxl.workbook.workbook.Worksheet, new_sheet: pyxl.workbook.workbook.Worksheet):
    """Map column labels used in new sheet to the main sheet positions"""

    label_positions = {}

    col = 1
    while main_sheet.cell(1, col).value is not None:
        main_cell = main_sheet.cell(1, col)
        new_col = col

        # Find new sheet position of current column label being examined in the main sheet
        while new_sheet.cell(1, new_col).value is not None:
            new_cell = new_sheet.cell(1, new_col)

            # Keep track of where the label exists in each
            if main_cell.value.lower() == new_cell.value.lower():
                label_positions[main_cell.value] = _LabelColumns(main_cell.column, new_cell.column)
                break

            new_col += 1

        # Column label in main does not exist in new, and should be ignored later on
        if main_cell.value not in label_positions:
            label_positions[main_cell.value] = _LabelColumns(main_cell.column, None)

        col += 1

    return label_positions
