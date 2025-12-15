@echo off
call venv\Scripts\activate.bat
python "PATH_merge_pdfs.py"
call venv\Scripts\deactivate.bat
exit
