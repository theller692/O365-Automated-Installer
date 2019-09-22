@echo off

wmic path SoftwareLicensingService get OA3xOriginalProductKey | findstr /v OA3xOriginalProductKey >%Temp%\productkey.txt

set /P VAR=<%Temp%\productkey.txt
del %temp%\productkey.txt

slmgr.vbs /ipk %VAR%

exit /b %errorlevel%