@echo off
REM Generate the ready-to-use Excel file with VBA macro code

echo Generating Excel file with VBA macro code...
python create_excel_with_macro_code.py

echo.
echo File created: Threshold_Explorer_Ready_to_Use.xlsx
echo.
echo This file contains the VBA macro code ready to copy-paste into Excel.
echo No need to open separate .bas files!
echo.
pause
