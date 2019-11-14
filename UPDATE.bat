@ECHO OFF
set RESPATH=%~sdp0
path %RESPATH%\git\bin\;%PATH%
Python36_64\python.exe update.py
IF %ERRORLEVEL% EQU 102 (
	UPDATE.bat
	GOTO EOF
)
PAUSE
:EOF