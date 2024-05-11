:: Reference: https://superuser.com/a/1097805

@ECHO OFF

SET EXEName=MSACCESS.exe
SET FileFullPath=C:\UDAM\finch_production_db_COM2.mdb

TASKLIST | FINDSTR /I "%EXEName%"
IF ERRORLEVEL 1 GOTO :StartDB
GOTO :EOF

:StartDB
::START "" "%FileFullPath%"
START "" /min "C:\Program Files\Microsoft Office\Office14\%EXEName%" "%FileFullPath%"
::START "" /max "C:\Program Files\Microsoft Office\Office14\%EXEName%" "%FileFullPath%"
GOTO :EOF