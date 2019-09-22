@Echo off
Title Autorun.bat
Echo Administrator account will be required
Echo Press any key to acknowledge . . .

REM Find the first usb drive letter and put to a file
wmic logicaldisk where drivetype=2 get deviceid |findstr /v "DeviceID" >%temp%\driveletter.txt

REM Set drive letter as variable for later use
set /P driveletter=<%temp%\driveletter.txt

del %temp%\driveletter.txt
cd /d %driveletter%\

pause > nul

REM Checking windows version. If not 8.1 or 10 install will create another file to run the install after restart.
systeminfo | find /i "Windows 10"
if not errorlevel 1 (
REM Win10 positive
   goto continue
)
if errorlevel 1 (
   goto checkwin8
)



:checkwin8

systeminfo | find /i "Windows 8.1"
if not errorlevel 1 (
REM Win8 positive
   goto continue
)
if errorlevel 1 (
REM Not Win8 or Win10, restart required
   goto restart
)

:continue

cd UninstallOffice_Ver3
REM This will run each vbscript on silent. 
REM Only the appropriate VBscript will actually uninstall items.
for %%i in (*.vbs) do cscript %%i ALL /NoCancel /Bypass 1

taskkill /f /im communicator.exe /t >nul
taskkill /f /im ucmapi.exe /t >nul
msiexec /x {81BE0B17-563B-45D4-B198-5721E6C665CD} /passive

cd..\o365 Installers\32
start Office365x86.bat


Echo Installing and opening Microsoft Teams
cd..\Teams x64
Teams_windows_x64.exe


REM Copying files to Ivanti Packages folder
cd ..\..\
mkdir "%programfiles(x86)%\LANDesk\LDClient\sdmcache\packages$"
robocopy  . "%programfiles(x86)%\LANDesk\LDClient\sdmcache\packages$" /copy:DAT /E /xf %driveletter%\InstallOffice.bat

pause

exit

:restart


taskkill /f /im communicator.exe /t >nul
taskkill /f /im ucmapi.exe /t >nul
msiexec /x {81BE0B17-563B-45D4-B198-5721E6C665CD} /passive

cd UninstallOffice_Ver3
REM This will run each vbscript on silent. 
REM Only the appropriate VBscript will actually uninstall items.
for %%i in (*.vbs) do cscript %%i ALL /NoCancel /Bypass 1

cd..\

echo cd /d %driveletter%\ >> Resume.bat
Echo cd o365 Installers\32 >> Resume.bat
Echo start Office365x86.bat >> Resume.bat
Echo cd..\Teams x64>> Resume.bat
Echo Teams_windows_x64.exe >> Resume.bat
Echo cd..\ >> Resume.bat
Echo del Resume.bat >> Resume.bat

shutdown -r -f -t 10 -c "Windows needs to restart"

