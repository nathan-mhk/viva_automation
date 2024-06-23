# Healthcare Report Templates

Simple macros that extend the table for production and assembly metrics for entering new data. Conditional formatting, Excel formulae, and macros were created to improve the efficiency of the data-entry process, and provide an easy-to-read summary for end-user.

Additional Excel file and VBA module were created to further improve the efficiency of data entry. Raw data were entered directly into `RawData.xlsx`. `VHC_Stick_MoveColumns.bas` were executed within each report file to facilitate the copying-and-pasting process of raw data.

Update 2024-06-23: To improve the overall efficiency of data entry, modifications, and verifications, an additional Excel file `RawData.xlsx` were created. This file contains a table that is the simplified and rearranged version of the `Prod` table in all of the stick workbooks. Raw data are adjusted as they were being entered into this file. To make it easier to copy from `RawData.xlsx` and paste into each stick workbook, an additional `MoveColumns` VBA module were implemented.
