@echo off
%~d0
cd %~dp0
SET PYTHONPATH=%cd%\python
SET PATH=%PATH%%cd%\python;%cd%\python\Scripts;
START pythonw.exe main.py
REM PAUSE
