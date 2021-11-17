%~d0
CD %~dp0
SET PYTHONPATH=%cd%\python;
SET PATH=%PATH%%cd%\python;
SET PATH=%PATH%%cd%\python\Scripts;
python.exe main.py
pause
