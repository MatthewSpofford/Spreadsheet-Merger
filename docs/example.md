# Example Use Case

Here is an example of how the merger tool would work against two sheets. In this example, these sheets are being used to
keep track of financial opportunities from different companies. The "ID Column" in this case will be `Opportunity ID`
(the default of the Spreadsheet Merger tool):

*Orig_sheet.xlsx*:

| Opportunity ID | Oppor. Lead  | Value | Notes                       |
|----------------|--------------|-------|-----------------------------|
| 1              | John Adams   | $100K | Follow up with John         |
| 2              | Kerry Burns  | $250K |                             |
| 3              | Liam Clam    | $80K  | What is the status of this? |
| 4              | Mark Daniels | $500K |                             |

*Appending_sheet.xlsx*:

| Opportunity ID | Oppor. Lead | Value |
|----------------|-------------|-------|
| 1              | John Adams  | $100K |
| 2              | Steve Pauls | $300K |
| 3              | Ryan Kraft  | $450K |
| 5              | Nancy Evans | $200K |

If you ran these tables through the merger tool, it would first scan through the "ID columns" of each sheet, which in
this case is the `Opportunity ID` column. It will then learn the following information:
* IDs 1, 2, and 3 exist in both sheets, and the columns should be updated
* ID 5 is new for the original sheet
* ID 4 does not exist in the appending sheet, and will be left untouched

With this information, it will merge will result into this new sheet:

| Opportunity ID | Oppor. Lead  | Value | Notes                       |
|----------------|--------------|-------|-----------------------------|
| 5              | Nancy Evans  | $200K |                             |
| 1              | John Adams   | $100K | Follow up with John         |
| 2              | Steve Pauls  | $300K |                             |
| 3              | Ryan Kraft   | $450K | What is the status of this? |
| 4              | Mark Daniels | $500K |                             |

Due to the smart merging process, this retains columns such as `Notes`, while also updating the columns `Oppor. Lead`
and `Value`.