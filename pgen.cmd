%~d0
cd %~dp0
set PATH=%PATH%%cd%;
python.exe %~dp0main.py
pause
