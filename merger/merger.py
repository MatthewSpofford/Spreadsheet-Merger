import openpyxl as pyxl


def merge_spreadsheets(main_file_path, new_file_path):
    main_sheet = pyxl.load_workbook(main_file_path)["Manager_Opportunity_Dashboard"]
    new_sheet = pyxl.load_workbook(new_file_path, read_only=True)["Manager_Opportunity_Dashboard"]

    # Map column labels used in new sheet to the main sheet positions

    # Find all "Opportunity ID's" in the main sheet

    # Update all previous opportunities in the main sheet with new sheet data

    # Insert new opportunities from at the top of the main sheet
