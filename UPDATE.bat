@echo off
pypy2-v5.9.0-win32\pypy.exe update.py
if %ERRORLEVEL% EQU 102 (
	UPDATE.bat
)
pause