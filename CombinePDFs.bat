@echo off
p:
set dir=P:\Donlin Recano
echo[
echo[
echo            DONLIN RECANIO - COMBINE PDFS 
echo ----------------------------------------------------
echo[
echo Before proceeding, ensure the steps below are completed:
echo[
echo   1. Create job folder at:
echo[
echo      "P:\Donlin Recano\--DS Job Num--\--DRC Job Num--" or
echo[
echo      "P:\Donlin Recano\Jose\Generic\--DRC Job Num--" (if there's no job ticket)
echo[
echo[
echo   2. In the job folder, create a folder named "Document". Place the PDFs to combine in this folder.
echo[
echo[
set /p startorquit=Type "y" when ready to process... 
if %startorquit%==y (
    goto start
) 
if %startorquit%==Y (
    goto start
) 
if not %startorquit%==Y (
    if not %startorquit%==y (
    goto EOF
    )
) 


:start
echo[
echo[
echo[
echo ----------------------------------------------------
set /p hasdsnumber="1.  DS Job Number (ds) or Jose Generic Folder (g): "
echo[
if %hasdsnumber%==ds (goto dsjobno)
if %hasdsnumber%==g (goto generic)
if %hasdsnumber%==DS (goto dsjobno)
if %hasdsnumber%==G (goto generic)

if not %hasdsnumber%==ds (
    if not %hasdsnumber%==DS (
    goto EOF
    )
) 
if not %hasdsnumber%==g (
    if not %hasdsnumber%==G (
    goto EOF
    )
) 

:dsjobno
set /p dsnumber="2.  Type DS Job Number: "
echo[
set /p drcnumber="3.  Type DRC Job Number: "
echo[
set folderpath=%dir%\%dsnumber%\%drcnumber%\Document
goto process

:generic
set /p drcnumber="2.  Type DRC Job Number: "
echo[
set folderpath=%dir%\Jose\Gen\%drcnumber%\Document
goto process

:process

cd %folderpath%
python "%dir%\script\Merge_And_List_PDFs.py" .
rename combined.pdf %drcnumber%_combined.pdf

:EOF
pause
