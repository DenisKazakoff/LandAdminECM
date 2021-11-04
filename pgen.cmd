%~d0
cd %~dp0
set PATH=%PATH%%cd%;
pythonw.exe %~dp0main.py
pause
