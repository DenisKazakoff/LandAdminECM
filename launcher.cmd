@echo off
%~d0
cd %~dp0

if exist %cd%\python\ (
	set PYTHONPATH=%cd%\python;
	set PATH=%PATH%%cd%\python;
	set PATH=%PATH%%cd%\python\Scripts;
)

pythonw.exe main.py
REM pause
