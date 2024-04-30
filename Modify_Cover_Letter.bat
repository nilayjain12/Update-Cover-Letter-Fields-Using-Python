@echo off
set /p position_name=Enter position name:
set /p company_name=Enter company name:
cd /d "C:\Users\njain\OneDrive - Cal State Fullerton\SPRING 2024\Nilay Jain Resume\Cover Letter"
python Python_Script_Change_PositionCompanyName.py "%position_name%" "%company_name%"
pause
