@echo off
setlocal enabledelayedexpansion

set "destinationFolder=C:\Users\User\Downloads\DailyReport\"

del /q "!destinationFolder!*"

@REM Declare an array constant called `productTypes`
set "productTypes[0]=DEO"
set "productTypes[1]=DER"
set "productTypes[2]=DSTF"
set "productTypes[3]=JAR"
set "productTypes[4]=MEO"

@REM set tgtYear=2023
@REM set tgtMonth=05

for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set "dt=%%I"
set /a tgtYear=%dt:~0,4%
set /a tgtMonth=%dt:~4,2%
set /a day=%dt:~6,2%
set /a day-=2
if %day% lss 1 (
    set /a tgtMonth-=1
    if !tgtMonth! lss 1 (
        set /a tgtMonth=12
        set /a tgtYear-=1
    )
    for %%I in (1 3 5 7 8 10 12) do if !tgtMonth!==%%I set day=31
    for %%I in (4 6 9 11) do if !tgtMonth!==%%I set day=30
    if !tgtMonth!==2 (
        set day=28
        set /a leapyear=tgtYear %% 4
        if !leapyear!==0 (
            set day=29
            set /a leapyear=tgtYear %% 100
            if !leapyear!==0 (
                set day=28
                set /a leapyear=tgtYear %% 400
                if !leapyear!==0 set day=29
            )
        )
    )
)
if %tgtMonth% lss 10 set tgtMonth=0%tgtMonth%

@REM echo Year: %tgtYear%
@REM echo Month: %tgtMonth%

@REM For each `productType`, copy the file into "C:\Users\User\Downloads\DailyReport"
for %%i in (0 1 2 3 4) do (
  set "productType=!productTypes[%%i]!"
  set sourceFile=Path\To\File\Finch !productType! Report\%tgtYear%\Finch !productType! Report %tgtYear%-%tgtMonth%.xlsm
  @REM echo %sourceFile%
  copy /y "!sourceFile!" "!destinationFolder!"
)

endlocal
