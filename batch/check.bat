@echo off
call F:\venv\Scripts\activate.bat
python "PATH_check_missing_pdfs.py"
call F:\venv\Scripts\deactivate.bat
pause
