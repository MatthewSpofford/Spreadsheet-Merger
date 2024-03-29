# Spreadsheet Merger

This app is used for merging two Excel spreadsheets into one spreadsheet based on an identifier column. 

This tool uses the library [openpyxl](https://pypi.org/project/openpyxl/3.0.10/) for opening, saving, and joining Excel
documents in Python.

## Installation

To install the latest version of this application, click the following link (currently only supports Windows devices):

**[Spreadsheet-Merger.zip](https://github.com/MatthewSpofford/Spreadsheet-Merger/releases/latest/download/Spreadsheet-Merger.zip)**

Once you have installed the zip file, you can unzip the file which will create a new folder named `Spreadsheet-Merger`.
This new folder will contain the entire application. You may move it wherever you desire.

If you would like to install an older version, go to the repositories [releases page](https://github.com/MatthewSpofford/Spreadsheet-Merger/releases)
and then click on the desired version. From there scroll down to the `Assests` section, and install the 
`Spreadsheet-Merger.zip`.

## How it works

This app is used for merging two Excel spreadsheets into one spreadsheet, where each sheet contains an "ID column" by
the same name for uniquely identifying each row. One spreadsheet is selected to be the "original" spreadsheet, while
the other is the "appending" sheet. The tool then goes through the ID column of both sheets.

Between the original and appending sheets, if the tool finds two rows with the same ID, it assumes that the appending
sheet has the most up-to-date columns for that ID. So for every column with the same name, it updates the contents of
the original row to be the same as the appending row. Since the tool uses column names to identify columns, it will even be
capable of updating rows if the columns are in a different order between sheets. Also, if there are any columns in the
original sheet that don't exist in the appending sheet, these are left unchanged in the original.

On top of that, if row IDs exist in the appending sheet that don't exist in the original sheet, the entire row will be
added to the top of the original sheet. Conversely, there are row IDs that exist in the original sheet that do not exist
in the appending sheet, then these rows will be left unchanged.

For an example of how to use the spreadsheet merger tool, check out this [example use case](docs/example.md).

## How to use it

1. In the installed `Spreadsheet-Merger`, find and run the `Spreadsheet-Merger.exe`.
3. Specify the original spreadsheet file to be merged against in the `Original Spreadsheet` textbox.
4. Specify the new spreadsheet file that will be merged in the `Appending Spreadsheet` textbox. 
5. Specify the "ID Column", which is used to uniquely identify each row. Defaults to "Opportunity ID".
6. Optionally, you can decide whether to replace the original spreadsheet after the merging process by checking the
`Replace original spreadsheet` checkbox. This checkbox is checked by default, meaning that the original spreadsheet
will be replaced.
   1. If you do *not* want to replace the original spreadsheet, you will need to specify the name of the new
   spreadsheet in the `Merged Spreadsheet Name` textbox.
7. Once you've entered the desired settings, click the `Merge` button to begin the spreadsheet merging process.
8. While the merging is completing, a progress bar will appear to indicate the status of the merge.
   1. If you would like to cancel the merge at any point, press the `Cancel` button.

Once you run the application for the first time, a configuration file is created in your home directory
(`C:\Users\<USER-PROFILE>\spreadsheet-merger.ini`). This enables you to keep your settings between each run of the tool. 

## Support

To support this application you may go to the [issues page](https://github.com/MatthewSpofford/Spreadsheet-Merger/issues/new/choose)
to report any bugs, or request a new feature.
