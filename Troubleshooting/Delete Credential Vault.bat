@echo off

REM Lists all credentials in the Windows Credential Manager vault
cmdkey /list | find /i "Microsoft"> "%TEMP%\creds.txt"

REM Reads the name of the actual Office 365 credential and instructs cmdkey to remove it.
FOR /F "tokens=2 delims= " %%G IN (%TEMP%\creds.txt) DO cmdkey /delete:%%G

REM Removes this temp file as it's not needed anymore
del "%TEMP%\creds.txt"

pause