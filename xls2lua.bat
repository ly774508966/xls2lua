echo off

:Start

python xls2lua.py test.xlsx TEST_SHEET

echo ===================================
echo xls2lua execute Success
echo ===================================

:Exit

pause
