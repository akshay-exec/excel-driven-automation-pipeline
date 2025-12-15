@echo off
call venv\Scripts\activate.bat
python "PATH_split_pdfs.py"
call venv\Scripts\deactivate.bat
pause
