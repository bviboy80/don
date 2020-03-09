@echo off
setlocal enabledelayedexpansion

set dir=\\dpfile.amstock.com\dpclients\Donlin Recano

:instructions
echo[
echo[
echo[
echo        DONLIN RECANIO LABELS PROCESSING - LEADER SHEET
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
echo 2. In the job folder, create a folder named "Data". Place data file to 
echo    process in this folder.
echo[
echo 3. Convert the Excel data files to CSV files.
echo[
echo[
set /p startorquit=Ready to process (y or n)... 
if %startorquit%==y (
    goto start) 
if %startorquit%==Y (
    goto start) 
if not %startorquit%==Y (
    if not %startorquit%==y (
    goto EOF
    )
) 
echo[
echo[
::Menu Options
:start
cls
echo[
echo[
echo ------------------------------------------
echo               Menu Options
echo ------------------------------------------
echo "back = back one, start = start"
echo[
echo[
echo Start
set /p hasjobnumber="DS job number (ds) or Generic folder (g): "
echo[
if %hasjobnumber%==ds (goto ds1)
if %hasjobnumber%==g (goto ds2)
if %hasjobnumber%==back (goto start)
if %hasjobnumber%==start (goto start)

:ds1
echo[
echo[
echo 1.
echo[
set /p dsjobno="DS job number: "
if %dsjobno%==back (goto start)
if %dsjobno%==start (goto start)

:ds2
echo[
echo[
if %hasjobnumber%==ds ( echo 2.
) else (echo 1.)
echo[
set /p drcjobno="DRC job number: "

if %drcjobno%==back (
   if %hasjobnumber%==ds (
       goto ds1
   ) else ( goto start )
)

if %drcjobno%==start (goto start)

:ds3
echo[
echo[
if %hasjobnumber%==ds ( echo 3.
) else (echo 2.)
echo[
set /p compname=Company Name (e.g., Rupari, Gander): 
if "%compname%"==back (goto ds2)
if "%compname%"==start (goto start)

:ds4
echo[
echo[
if %hasjobnumber%==ds ( echo 4.
) else (echo 3.)
echo[
set /p location="Domestic (d) or Foreign (f): "
if %location%==d (set place=Domestic)
if %location%==f (set place=Foreign)
if %location%==back (goto ds3)
if %location%==start (goto start)

:ds5
echo[
echo[
if %hasjobnumber%==ds ( echo 5.
) else (echo 4.)
echo[
set /p data="Data file to process: "
if "%data%"==back (goto ds4)
if "%data%"==start (goto start)

:ds6
echo[
echo[
if %hasjobnumber%==ds ( echo 6.
) else (echo 5.)
echo[
set /p pdffile="PDF Document to process: "
if "%pdffile%"==back (goto ds5)
if "%pdffile%"==start (goto start)

:ds7
echo[
echo[
if %hasjobnumber%==ds ( echo 7.
) else (echo 6.)
echo[
echo 10    =  #10 Double Window Env
echo 6x9   =  6" x 9" DRC Window Env
echo 9x12  =  9" x 12" DRC Window Env
echo 9x12N =  9" x 12" DRC - NO RETURN ADDRESS
echo[
set /p envelope="Mailing Envelope(10, 6x9, 9x12, 9x12N): "
if !envelope!==6x9 (
   set pobox=0
   goto ds9
)
if !envelope!==9x12N (
   set pobox=0
   goto ds9
)
if !envelope!==back (goto ds6)
if !envelope!==start (goto start) 

:ds8
echo[
echo[
if %hasjobnumber%==ds ( echo 8.
) else ( echo 9.)
echo[
echo No.   PO Box            Address_Line,            City_State_Zip
echo[
echo 0.    NO RETURN ADDRESS
echo 1.    P.O. Box 899      Madison Square Station,  New York, NY 10010
echo 2     P.O. Box 2047     Murray Hill Station,     New York, NY 10156
echo 3.    P.O. Box 2062     Murray Hill Station,     New York, NY 10156
echo 4.    P.O. Box 2070     Murray Hill Station,     New York, NY 10156
echo 5.    P.O. Box 192016   Blythebourne Station,    Brooklyn, NY 11219
echo 6.    P.O. Box 192042   Blythebourne Station,    Brooklyn, NY 11219
echo 7.    P.O. Box 192328   Blythebourne Station,    Brooklyn, NY 11219
echo 8.    P.O. Box 199001   Blythebourne Station,    Brooklyn, NY 11219
echo 9.    P.O. Box 199012   Blythebourne Station,    Brooklyn, NY 11219
echo 10.   P.O. Box 199043   Blythebourne Station,    Brooklyn, NY 11219
echo[
set /p pobox="Select Return Address: "
if !pobox!==back ( goto ds7)
if !pobox!==0 ( set hasaddress=False )
if not !pobox!==0 ( set hasaddress=True )

set /p voting="Use -- Attn: Voting Department -- (y or n)"
if !voting!==y ( set votebool=True )
if !voting!==n ( set votebool=False )




:ds9
echo[
echo[
if !envelope!==6x9 ( set hasaddress=False ) 
if !envelope!==6x9 ( set votebool=True ) 
if !envelope!==9x12N ( set hasaddress=False ) 
if !envelope!==9x12N ( set votebool=True ) 


echo[
set /p recordsplit=Number of Records per group split: 
 
if %recordsplit%==start (goto start)

:ds10
echo[
echo[
set /p leadersheet="Print PDF with Leader Sheet (p) or Leader sheet only (l): "
if %leadersheet%==p (
    set output=Standard_Print
)
if %leadersheet%==l (
    set output=Standard_Leader
)
if %leadersheet%==back (goto ds9)
if %leadersheet%==start (goto start)


set /p leaderfrontorback="Leader Sheet in back (b) or front (f): "
if !leaderfrontorback!==b (
    set workflow=Leader Sheet Layout.wfd
)
if !leaderfrontorback!==f (
    set workflow=Leader Sheet Layout Leaderfront.wfd
)

echo[
echo[
echo[
:: Set Filepaths for processing
if %hasjobnumber%==ds ( 
    set folderpath=%dir%\%dsjobno%\%drcjobno%
    set filename=%dsjobno%_%compname%_%drcjobno%_%place%
) 
if %hasjobnumber%==g (
   set folderpath=%dir%\Jose\Gen\%drcjobno%
   set filename=%drcjobno%_%compname%__%place%
)
set logfile=!folderpath!\Data\%drcjobno%.log
set pdf=!folderpath!\Document\%pdffile%



"G:\PrintNet T Designer\PNetTC.exe" "%dir%\PrintNet\Shaun\!workflow!" ^
-difData "!folderpath!\Data\%data%" ^
-difPOaddress "%dir%\script\POboxes.csv" ^
-PoBoxParams "!pobox!" ^
-pdfDocumentParams "%pdf%" ^
-EnvelopeParams "%envelope%" ^
-RecordsPerGroupParams %recordsplit% ^
-HasReturnAddressParams !hasaddress! ^
-VoteDeptParams !votebool! ^
-o !output! ^
-c "P:\Donlin Recano\Printnet\Shaun\DRC_Leader_Sheet.job" ^
-pc OCE ^
-dc "DRC Leader Sheet" ^
-f "!folderpath!\Print\!filename!_%%g.ps" ^
-e AdobePostScript3 ^
-la "%logfile%" ^
-splitbygroup
goto EOF


endlocal
:EOF
pause
