@echo off
set /p week= Input which week you want to create report ^>
set /p line_of_week= Input which line week you just type ^>

echo Preparing create report for week %week% .....

python execute_ver2.py %week% %line_of_week%


echo ------------------------------------------Bye------------------------------------------
pause