@ECHO OFF
pypy2-v5.9.0-win32\pypy.exe update.py
IF %ERRORLEVEL% EQU 102 (
	UPDATE.bat
	GOTO EOF
)
PAUSE
:EOF