@echo off
%~d0
CD %~dp0
SET PYTHONPATH=%cd%\python;
SET PATH=%PATH%%cd%\python;
SET PATH=%PATH%%cd%\python\Scripts;
pythonw.exe main.py
REM pause