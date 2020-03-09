@echo off
setlocal enabledelayedexpansion

set dir=\\dpfile.amstock.com\dpclients\Donlin Recano

:instructions
echo[
echo[
echo[
echo        DONLIN RECANIO LABELS PROCESSING - 10 UP - CASELINE DATA
echo ____________________________________________________
echo[
echo Before proceeding, ensure the steps below are completed:
echo[
echo 1. Got to --^> "P:\Donlin Recano". ^If there is a DS job ticket, 
echo    Ensure there is a folder labeled with DS job number. Inside the DS
echo    job folder should be a susequent folder labeled with the DRC job number
echo[
echo    ^If there is NO DS job ticket, Got to --^> "P:\Donlin Recano\Jose\Gen"
echo    Inside the "Gen" folder, ensure there is a folder labeled with DRC job number.
echo[
echo 2. In the job folder, create a folder named "Data". Place data file to process in this folder.
echo[
echo 3. Convert the Excel data files to CSV files.
echo[
echo[
set /p startorquit=Ready to process (y or n)... 
if %startorquit%==y (goto start)
if %startorquit%==n (goto EOF)
echo[
echo[
::Menu Options
:start
echo[
echo[
echo ------------------------------------------
echo               Menu Options
echo ------------------------------------------
echo "back = back one, quit = quit, start = start"
echo[
echo[
echo Start
set /p hasjobnumber="DS job number (ds) or Generic folder (g): "
echo[
if %hasjobnumber%==ds (goto ds1)
if %hasjobnumber%==g (goto ds2)
if %hasjobnumber%==back (goto start)
if %hasjobnumber%==quit (goto EOF)
if %hasjobnumber%==start (goto start)
if %hasjobnumber%==DS (goto ds1)
if %hasjobnumber%==G (goto ds2)
if %hasjobnumber%==BACK (goto start)
if %hasjobnumber%==QUIT (goto EOF)
if %hasjobnumber%==START (goto start)
echo[

set workflow=10UP-DRC Address.wfd
set labeltemp=10UP


:ds1
echo[
echo 1 of 5
set /p dsjobno="DS job number: "
if %dsjobno%==back (goto start)
if %dsjobno%==quit (goto EOF)
if %dsjobno%==start (goto start)
if %dsjobno%==BACK (goto start)
if %dsjobno%==QUIT (goto EOF)
if %dsjobno%==START (goto start)
:ds2
echo[
echo[
if %hasjobnumber%==ds( echo 2 of 5
) else ( echo 1 of 4)
echo[
set /p drcjobno="DRC job number: "
if %drcjobno%==back (
    if %hasjobnumber%==ds ( goto ds1
    )
    if %hasjobnumber%==g ( goto start
    )
)
if %drcjobno%==BACK (
    if %hasjobnumber%==ds ( goto ds1
    )
    if %hasjobnumber%==g ( goto start
    )
)
if %drcjobno%==quit (goto EOF)
if %drcjobno%==start (goto start)

if %drcjobno%==QUIT (goto EOF)
if %drcjobno%==START (goto start)
:ds3
echo[
echo[
if %hasjobnumber%==ds( echo 3 of 5
) else ( echo 2 of 4)
echo[
echo 3 of 5
set /p compname="Company Name: "
if "%compname%"==back (goto ds2)
if "%compname%"==quit (goto EOF)
if "%compname%"==start (goto start)
if "%compname%"==BACK (goto ds2)
if "%compname%"==QUIT (goto EOF)
if "%compname%"==START (goto start)
:ds4
echo[
echo[
if %hasjobnumber%==ds( echo 4 of 5
) else ( echo 3 of 4)
echo[
echo 4 of 5
set /p location="Domestic (d) or Foreign (f): "
if %location%==d (set place=DOMESTIC)
if %location%==f (set place=FOREIGN)
if %location%==back (goto ds3)
if %location%==quit (goto EOF)
if %location%==start (goto start)
if %location%==D (set place=DOMESTIC)
if %location%==F (set place=FOREIGN)
if %location%==BACK (goto ds3)
if %location%==QUIT (goto EOF)
if %location%==START (goto start)
:ds5
echo[
echo[
if %hasjobnumber%==ds( echo 5 of 5
) else ( echo 4 of 4)
echo[
echo 5 of 5
set /p data="Data file to process: "
if "%data%"==back (goto ds4)
if "%data%"==quit (goto EOF)
if "%data%"==start (goto start)
if "%data%"==BACK (goto ds4)
if "%data%"==QUIT (goto EOF)
if "%data%"==START (goto start)

if %hasjobnumber%==ds (
    set folderpath=%dir%\%dsjobno%\%drcjobno%
    set filename=%dsjobno%_%compname%_%drcjobno%_%place%
)
if %hasjobnumber%==g (
    set folderpath=%dir%\Jose\Gen\%drcjobno%
    set filename=%drcjobno%_%compname%_%place%
)


:label
echo[
set workflow=10UP-DRC Address.wfd
set labeltemp=10UP
set logfile=!folderpath!\Data\%drcjobno%.log
"G:\PrintNet T Designer\PNetTC.exe" "%dir%\PrintNet\Shaun\%workflow%" ^
-difData "!folderpath!\Data\%data%" ^
-JobNameParams "!filename!_%labeltemp%" ^
-o "Labels-Portrait" ^
-f "!folderpath!\Labels\!filename!_%labeltemp%.ps" ^
-e AdobePostScript3 ^
-la "%logfile%"
goto EOF

endlocal
:EOF
pause
