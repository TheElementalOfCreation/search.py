@echo off
set CONVPATH="%~dp0ImageToPDF\img2pdf.exe"
set RESPATH=%~sdp0
set PYPYPATH=%~sdp0pypy2-v5.9.0-win32\pypy.exe
set MSG=%~sdp0msg.py
if '%1' == '' (
	goto MSG
)
if not exist %1 (
	goto DNE
)
if exist %1\* (
	for %%i in (%1\*.tif*) do ( %CONVPATH% "%%i" --output "%%~dpni.pdf" )
	goto MSG
)
for %%g in (".tiff" ".tif" ".TIF" ".TIFF") do if "%~x1" == %%g (
	%CONVPATH% %1 --output "%~dpn1.pdf"
)

:MSG

%PYPYPATH% %MSG% "%~s1"
pause
goto EOF

:DNE

echo The file or directory "%1" does not exist or cannot be found.
echo If this is an error, please report it to the developer.
pause

:EOF