@echo off

REM cscript.exe will load the details about the Office 365 Product Key
REM findstr Last finds the line starting with "Last"
REM All output is going to a temp file in %TEMP%
cscript.exe "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /dstatus | findstr Last>> "%TEMP%\List.txt"
cscript.exe "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /dstatus | findstr Last>> "%TEMP%\List.txt"
REM Setting the text file as a variable to read from
Set /p var=<%TEMP%\List.txt

REM Reading the 8th "word" (determined by spaces) from each line of the text file (this grabs only the 5 digits needed)
REM Then outputting to a second temporary file
for /F "tokens=8" %%A in (%TEMP%\List.txt) do echo %%A>> %TEMP%\Key.txt

REM Removing the first temp file as it's not needed anymore
del %Temp%\List.txt

REM Reads each line (each line is a key) from the temp file and for every time there's a key attempts to uninstall the key
For /F %%B in (%TEMP%\Key.txt) do cscript.exe "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /unpkey:%%B
For /F %%B in (%TEMP%\Key.txt) do cscript.exe "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /unpkey:%%B

REM Removes this temp file as it's not needed anymore
del %temp%\Key.txt

REM Lists all credentials in the Windows Credential Manager vault
cmdkey /list | find /i "Microsoft"> "%TEMP%\creds.txt"

REM Reads the name of the actual Office 365 credential and instructs cmdkey to remove it.
FOR /F "tokens=2 delims= " %%G IN (%TEMP%\creds.txt) DO cmdkey /delete:%%G

REM Removes this temp file as it's not needed anymore
del "%TEMP%\creds.txt"