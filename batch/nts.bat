@echo off
call venv\Scripts\activate.bat
python "PATH_main.py" NTS
call venv\Scripts\deactivate.bat
exit
