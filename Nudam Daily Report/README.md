# Nudam Daily Report

Automation Scripts and Files for Nudam-to-PC data logging.

`Form_FormRun.bas` & `Form_frmLineInfo.bas` - The main scripts that handles the real-time data logging between 住友 (Sumitomo) injection molding machines and the computer.
> Disclaimer: Most of the code within these two `.bas` files were written by someone else. My task is to make it work again, and add any additional improvements if necessary.

The Microsoft Access database is designed to exit every 20 minutes. A task was scheduled in the Task Scheduler to automatically reopen the database file after it was exited.

`RefreshLineDetails.bas` were executed every time the Access database was launched, to ensure the production line details on the database matches the information on the actual machine.

Due to hardware limitations, some values of the shift-counter has to be manually adjusted. Adjustments and changes are all made and applied respectively within the Access database.

`ExportDaily.bas` were executed once every hour to generate a daily report that consolidates the production metrics of the previous day and the current shift.

Update 2024-05-10: As an improvement to suit the needs and requirements for the current state of the workflow, `ExportDaily.bas` and the Power Queries inside `Finch Plant Daily Production Summary.xlsx` can now handle the report generation of specific dates.

Update 2024-06-23: To improve the overall efficiency of data entry, modifications, and verifications, the use and functionalities of `ManualUpdate.xlsm` were refactored & integrated into various SQL queries and macros within the Access database.
