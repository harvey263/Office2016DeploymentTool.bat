cls
@echo off
::HarveyTDixon2016
title Office 2016 Deployment Tool
::https://www.microsoft.com/en-us/download/details.aspx?id=49117
::-------------------------------------------------------------------------------
"%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system" >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (echo Elevating... & GOTO UAC1) else (GOTO UAC2)
:UAC1
echo SET UAC = CreateObject^("Shell.Application"^) > "%TEMP%\uac.vbs"
echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%TEMP%\uac.vbs"
"%TEMP%\uac.vbs"
del /q /f "%TEMP%\uac.vbs" & exit /b
:UAC2
IF EXIST "%TEMP%\uac.vbs" (del /q /f "%TEMP%\uac.vbs")
pushd %CD% & CD /d %~dp0
::-------------------------------------------------------------------------------
::::::::::::::::::::::::::::::
::>    -Required Access-   <::
::> c$ Share               <::
::> Remote Registry        <::
::> Remote Scheduled Tasks <::
::::::::::::::::::::::::::::::

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::> Create a shared folder your clients have access to <::
::> Put the Office Deployment Tool setup.exe inside it <::
::> Change the below UNC path to your shared directory <::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

SET "DEPLOYDIR=\\SERVER01\dfs$\InstallOffice"

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::> If you are not in a domain, remove the 4 instances of %USERDOMAIN%\ <::
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:STARTODT2016
cls
@echo off
echo ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
echo ³Office 2016 Deployment Tool³
echo ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
echo.
ENDLOCAL
SETLOCAL ENABLEDELAYEDEXPANSION
SET "XMLDIR=%DEPLOYDIR%\XML"
SET "LOGDIR=%DEPLOYDIR%\Logs"
SET "SETUPDIRLOC=C:\Windows\Temp\Office2016Setup"
SET "SETUPDIRNET=c$\Windows\Temp\Office2016Setup"
SET "CURDATE=%date:~-4,4%%date:~-10,2%%date:~-7,2%"
SET "OC2RCLIENT=C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
SET "KEYC2ROFFICE16=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us"
SET "KEYC2RPROJECT16=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ProjectStdXVolume - en-us"
SET "KEYC2RVISIO16=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VisioStdXVolume - en-us"
SET "KEYC2RPOLICY=HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate"
SET "KEYC2RCONFIG=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
IF NOT EXIST "%DEPLOYDIR%\Download" MD "%DEPLOYDIR%\Download" 2>nul
FOR /d %%R in ("%LOGDIR%\Temp\*.*") do RD "%%R" /s /q 2>nul
IF NOT EXIST "%LOGDIR%\Temp" MD "%LOGDIR%\Temp" 2>nul
IF NOT EXIST "%XMLDIR%" MD "%XMLDIR%" 2>nul
DEL /f /q "%LOGDIR%\Temp\*.*" >nul 2>nul
SET "VALUEVERC2R=VersionToReport"
SET "VALUEVERSION=DisplayVersion"
SET "VALUEUPDATEPATH=updatepath"
SET "VALUEUPDATEURL=UpdateUrl"
SET "VALUEPLATFORM=Platform"
SET "VALUENAME=DisplayName"
SET "REGADDINSTALLC2R=N"
SET "SETUPSOURCE="
SET "SETUPTYPE="
SET "BITNESS="
echo [I] Install
echo [U] Uninstall
echo [C] Check for updates
echo [D] Download updates
echo [V] Version history
echo [X] Exit
echo.
choice /c IUCDVX /N /M ">"
IF ERRORLEVEL 6 exit
IF ERRORLEVEL 5 GOTO VERSIONC2R
IF ERRORLEVEL 4 GOTO DOWNLOADC2R
IF ERRORLEVEL 3 GOTO UPDATEC2R
IF ERRORLEVEL 2 GOTO UNINSTALLC2R
IF ERRORLEVEL 1 GOTO INSTALLC2R

::~~~~~~~~~~~~~~~::
::Version history::
::~~~~~~~~~~~~~~~::
:VERSIONC2R
echo.
START "" "%programfiles%\Internet Explorer\iexplore.exe" -noframemerging "https://technet.microsoft.com/en-us/library/mt592918.aspx"
GOTO STARTODT2016

::~~~~~~~~~~~~~~~~::
::Download updates::
::~~~~~~~~~~~~~~~~::
:DOWNLOADC2R
echo.
echo ^> This will download the selected channel to:
echo ^> "%DEPLOYDIR%\Download"
echo ^> See version history for what releases are currently available.
echo.
echo [C] Current channel
echo [D] Deferred channel
echo [F] First release for deferred channel
echo [X] Cancel
echo.
choice /c CDFX /N /M ">"
IF ERRORLEVEL 4 GOTO STARTODT2016
IF ERRORLEVEL 3 SET "C2RCHANNEL=FirstReleaseDeferred" & GOTO RUNDOWNLOADC2R
IF ERRORLEVEL 2 SET "C2RCHANNEL=Deferred" & GOTO RUNDOWNLOADC2R
IF ERRORLEVEL 1 SET "C2RCHANNEL=Current" & GOTO RUNDOWNLOADC2R

:RUNDOWNLOADC2R
echo.
SET "PRODUCT_ID=O365ProPlusRetail"
SET "SETUPTYPE=Download"
call :SETBITNESSC2R
call :XMLDOWNLOADC2R
echo.
echo Downloading, please wait...
"%DEPLOYDIR%\setup.exe" /download "%XMLDIR%\Config%SETUPTYPE%%PRODUCT_ID%-%C2RCHANNEL%_%BITNESS%.xml"
IF %ERRORLEVEL% EQU 0 (
echo. & echo Download SUCCESS^^!
) else (
echo. & echo Download FAILED^^!
)
echo.
pause
GOTO STARTODT2016

::~~~~~~~~~~~~~~~~~::
::Check for updates::
::~~~~~~~~~~~~~~~~~::
:UPDATEC2R
echo.
call :SETMACHINEC2R

:QUERYSOURCEC2R
SET "SETUPTYPE=Update"
call :REGQUERYC2R
IF %O365PRODUCTC2R% NEQ NUL (echo O365: Installed) else (echo O365: Not installed)
IF %VISIOPRODUCTC2R% NEQ NUL (echo Visio: Installed) else (echo Visio: Not installed)
IF %PROJECTPRODUCTC2R% NEQ NUL (echo Project: Installed) else (echo Project: Not installed)
IF %O365PRODUCTC2R% EQU NUL IF %VISIOPRODUCTC2R% EQU NUL IF %PROJECTPRODUCTC2R% EQU NUL (
echo. & echo ^> No Office products are currently installed^^! & echo. & pause & GOTO STARTODT2016
)
echo.
call :VERSIONREPORTC2R
SET "SOURCEVERSIONC2R="
FOR /f "tokens=*" %%U in ('dir /a:d /b "%UPDATEPATHC2R%\Office\Data" 2^>nul') do (SET "SOURCEVERSIONC2R=%%U")
IF NOT DEFINED SOURCEVERSIONC2R echo Update source no longer exists^^! & echo. & GOTO CHOICEPATHC2R

echo ^> Version at source
echo %SOURCEVERSIONC2R% %PLATFORMC2R%
echo.
IF %CLIENTVERSIONC2R% EQU %SOURCEVERSIONC2R% (
echo -----------------------
SET "VERSTATC2R=0" & echo ^> Client is up to date^^!
echo -----------------------
)
IF %CLIENTVERSIONC2R% LSS %SOURCEVERSIONC2R% (
echo ----------------------
SET "VERSTATC2R=1" & echo ^> Update is available^^!
echo ----------------------
)
IF %CLIENTVERSIONC2R% GTR %SOURCEVERSIONC2R% (
echo ---------------------------------------
SET "VERSTATC2R=2" & echo ^> Client version is higher than source^^!
echo ---------------------------------------
)

:CHOICEUPDATEC2R
echo.
echo [1] Install updates
echo [2] Change update path
echo [X] Cancel
echo.
choice /c 12X /N /M ">"
IF ERRORLEVEL 3 GOTO STARTODT2016
IF ERRORLEVEL 2 GOTO CHANGEPATHC2R
IF ERRORLEVEL 1 GOTO RUNUPDATEC2R

:CHOICEPATHC2R
choice /c YN /N /M "Do you want to change the update path? <Y/N>"
IF ERRORLEVEL 2 GOTO STARTODT2016
IF ERRORLEVEL 1 GOTO CHANGEPATHC2R

:CHANGEPATHC2R
echo.
call :SETSOURCEC2R
SET "NEWPATHVERSIONC2R="
SET "UPDATEPATHC2R=%DEPLOYDIR%\%SETUPSOURCE%\%PLATFORMC2R%"
FOR /f "tokens=*" %%T in ('dir /a:d /b "%UPDATEPATHC2R%\Office\Data" 2^>nul') do (SET "NEWPATHVERSIONC2R=%%T")
IF NOT DEFINED NEWPATHVERSIONC2R (
echo. & echo Error - no source files found at this path^^! & echo. & pause & echo. & GOTO QUERYSOURCEC2R
)
echo.
echo ^> Path will be changed to
echo %UPDATEPATHC2R%
echo.
echo ^> Version at source
echo %NEWPATHVERSIONC2R% %PLATFORMC2R%
echo.
IF %CLIENTVERSIONC2R% GTR %NEWPATHVERSIONC2R% (
echo ------------------------------------------------
echo ^> Client version is higher than source^^!
echo ^> Changing to this path will require a downgrade
echo ------------------------------------------------ & echo.
)
choice /c YN /N /M "Do you want to continue? <Y/N>"
IF ERRORLEVEL 2 GOTO QUERYSOURCEC2R
IF ERRORLEVEL 1 GOTO RUNCHANGEPATHC2R

:RUNCHANGEPATHC2R
IF %CLIENTVERSIONC2R% GTR %NEWPATHVERSIONC2R% (
SET "SKIPREGPATH=N" & call :DOWNGRADEC2R & call :REGQUERYC2R & GOTO QUERYSOURCEC2R
)
echo.
call :REGADDPATHC2R
echo. & pause & echo. & call :REGQUERYC2R & GOTO QUERYSOURCEC2R

:RUNUPDATEC2R
IF %VERSTATC2R% EQU 0 (
echo.
echo --------------------------------------------
echo ^> Client version already matches the source^^!
echo --------------------------------------------
echo. & pause & GOTO CHOICEUPDATEC2R
)
IF %VERSTATC2R% EQU 2 (
echo.
echo ---------------------------------------
echo ^> Client version is higher than source^^!
echo ---------------------------------------
echo.
choice /c YN /N /M "Do you want to downgrade the installation? <Y/N>"
IF ERRORLEVEL 2 GOTO QUERYSOURCEC2R
IF ERRORLEVEL 1 SET "SKIPREGPATH=Y" & call :DOWNGRADEC2R & call :REGQUERYC2R
GOTO QUERYSOURCEC2R
)
call :OC2RSWITCH1
call :DEPLOYSETUPC2R
IF NOT EXIST "\\%MACHINE%\%SETUPDIRNET%\Office2016_%SETUPTYPE%.vbs" echo Failed to create setup files on %MACHINE%^^! & echo. & pause & GOTO STARTODT2016
IF NOT EXIST "\\%MACHINE%\%SETUPDIRNET%\Office2016_%SETUPTYPE%.bat" echo Failed to create setup files on %MACHINE%^^! & echo. & pause & GOTO STARTODT2016
call :SETAUTHORC2R
call :SCHEDCREATEC2R
echo.
choice /c YN /N /M "Run update now? <Y/N>"
IF ERRORLEVEL 2 call :CLEANUPC2R & GOTO STARTODT2016
IF ERRORLEVEL 1 GOTO UPDATENOWC2R

:UPDATENOWC2R
call :SCHEDRUNC2R
call :UPDATEWAITC2R
echo. & pause & echo. & GOTO QUERYSOURCEC2R

::~~~~~~~~~::
::Uninstall::
::~~~~~~~~~::
:UNINSTALLC2R
echo.
SET "PRODUCT_ID="
SET "SETUPTYPE=Uninstall"
SET "REMOVEALLC2R=False"
SET "BITNESS=64x86"
call :SETMACHINEC2R

:STARTUNINSTALLC2R
call :REGQUERYC2R
IF %O365PRODUCTC2R% NEQ NUL (echo O365: Installed %O365VERSIONC2R% %PLATFORMC2R%) else (echo O365: Not installed)
IF %VISIOPRODUCTC2R% NEQ NUL (echo Visio: Installed %VISIOVERSIONC2R% %PLATFORMC2R%) else (echo Visio: Not installed)
IF %PROJECTPRODUCTC2R% NEQ NUL (echo Project: Installed %PROJECTVERSIONC2R% %PLATFORMC2R%) else (echo Project: Not installed)
IF %O365PRODUCTC2R% EQU NUL IF %VISIOPRODUCTC2R% EQU NUL IF %PROJECTPRODUCTC2R% EQU NUL (
echo. & echo ^> No Office products are currently installed^^! & echo. & pause & GOTO STARTODT2016
)
echo.
echo [A] All
echo [O] O365
echo [V] Visio
echo [P] Project
echo [X] Cancel
echo.
choice /c AOVPX /N /M ">"
echo.
IF ERRORLEVEL 5 GOTO STARTODT2016
IF ERRORLEVEL 4 SET "PRODUCT_ID=ProjectStdXVolume" & GOTO UNINSTALLPROJECTC2R
IF ERRORLEVEL 3 SET "PRODUCT_ID=VisioStdXVolume" & GOTO UNINSTALLVISIOC2R
IF ERRORLEVEL 2 SET "PRODUCT_ID=O365ProPlusRetail" & GOTO UNINSTALLO365C2R
IF ERRORLEVEL 1 SET "PRODUCT_ID=AllOfficeProducts" & SET "REMOVEALLC2R=True" & GOTO CONFIGUREUNINSTALLC2R

:UNINSTALLPROJECTC2R
IF %PROJECTPRODUCTC2R% EQU NUL (
echo ^> Project 2016 is not installed on %MACHINE%
echo. & pause & GOTO STARTODT2016
)
GOTO CONFIGUREUNINSTALLC2R

:UNINSTALLVISIOC2R
IF %VISIOPRODUCTC2R% EQU NUL (
echo ^> Visio 2016 is not installed on %MACHINE%
echo. & pause & GOTO STARTODT2016
)
GOTO CONFIGUREUNINSTALLC2R

:UNINSTALLO365C2R
IF %O365PRODUCTC2R% EQU NUL (
echo ^> Office 2016 is not installed on %MACHINE%
echo. & pause & GOTO STARTODT2016
)
GOTO CONFIGUREUNINSTALLC2R

:CONFIGUREUNINSTALLC2R
call :REGQUERYC2R
call :DEPLOYSETUPC2R
IF NOT EXIST "\\%MACHINE%\%SETUPDIRNET%\Office2016_%SETUPTYPE%.vbs" echo Failed to create setup files on %MACHINE%^^! & echo. & pause & GOTO STARTODT2016
IF NOT EXIST "\\%MACHINE%\%SETUPDIRNET%\Office2016_%SETUPTYPE%.bat" echo Failed to create setup files on %MACHINE%^^! & echo. & pause & GOTO STARTODT2016
call :XMLUNINSTALLC2R
call :SETAUTHORC2R
call :SCHEDCREATEC2R
echo.
echo You are about to uninstall %PRODUCT_ID% on %MACHINE%^^!
echo.
choice /c YN /N /M "Do you want to continue? <Y/N>"
IF ERRORLEVEL 2 call :CLEANUPC2R & GOTO STARTODT2016
IF ERRORLEVEL 1 GOTO RUNUNINSTC2R

:RUNUNINSTC2R
call :SCHEDRUNC2R
call :SETUPWAITC2R
echo. & pause & echo. & GOTO STARTUNINSTALLC2R

::~~~~~~~::
::Install::
::~~~~~~~::
:INSTALLC2R
echo.
SET "SETUPTYPE=Install"
call :SETMACHINEC2R

:STARTINSTALLC2R
call :REGQUERYC2R
IF %O365PRODUCTC2R% NEQ NUL (echo O365: Installed %O365VERSIONC2R% %PLATFORMC2R%) else (echo O365: Not installed)
IF %VISIOPRODUCTC2R% NEQ NUL (echo Visio: Installed %VISIOVERSIONC2R% %PLATFORMC2R%) else (echo Visio: Not installed)
IF %PROJECTPRODUCTC2R% NEQ NUL (echo Project: Installed %PROJECTVERSIONC2R% %PLATFORMC2R%) else (echo Project: Not installed)
IF %O365PRODUCTC2R% NEQ NUL IF %VISIOPRODUCTC2R% NEQ NUL IF %PROJECTPRODUCTC2R% NEQ NUL (
echo. & echo ^> All Office products are already installed^^! & echo. & pause & GOTO STARTODT2016
)
echo.
echo Select product to install:
echo.
echo [O] O365
echo [V] Visio
echo [P] Project
echo [X] Cancel
echo.
choice /c OVPX /N /M ">"
echo.
IF ERRORLEVEL 4 GOTO STARTODT2016
IF ERRORLEVEL 3 SET "PRODUCT_ID=ProjectStdXVolume" & GOTO INSTALLPROJECTC2R
IF ERRORLEVEL 2 SET "PRODUCT_ID=VisioStdXVolume" & GOTO INSTALLVISIOC2R
IF ERRORLEVEL 1 SET "PRODUCT_ID=O365ProPlusRetail" & GOTO INSTALLO365C2R

:INSTALLPROJECTC2R
IF %PROJECTPRODUCTC2R% NEQ NUL (
echo ^> Project 2016 is already installed on %MACHINE%
echo ^> Use the Uninstall or Update functions instead
echo. & pause & GOTO STARTODT2016
)
GOTO CONFIGUREINSTALLC2R

:INSTALLVISIOC2R
IF %VISIOPRODUCTC2R% NEQ NUL (
echo ^> Visio 2016 is already installed on %MACHINE%
echo ^> Use the Uninstall or Update functions instead
echo. & pause & GOTO STARTODT2016
)
GOTO CONFIGUREINSTALLC2R

:INSTALLO365C2R
IF %O365PRODUCTC2R% NEQ NUL (
echo ^> Office 2016 is already installed on %MACHINE%
echo ^> Use the Uninstall or Update functions instead
echo. & pause & GOTO STARTODT2016
)
GOTO CONFIGUREINSTALLC2R

:CONFIGUREINSTALLC2R
IF %PLATFORMC2R% EQU NUL (
call :SETBITNESSC2R & echo.
) else (
SET "BITNESS=%PLATFORMC2R%"
)
echo Select install/update source: & echo.
call :SETSOURCEC2R
SET "INSTALLVERSIONC2R="
SET "INSTALLPATHC2R=%DEPLOYDIR%\%SETUPSOURCE%\%BITNESS%"
FOR /f "tokens=*" %%I in ('dir /a:d /b "%INSTALLPATHC2R%\Office\Data" 2^>nul') do (SET "INSTALLVERSIONC2R=%%I")
IF NOT DEFINED INSTALLVERSIONC2R (
echo. & echo Error - no source files found at this path^^! & echo. & pause & echo. & GOTO STARTODT2016
)
echo.
IF %CLIENTVERSIONC2R% NEQ NUL IF %INSTALLVERSIONC2R% NEQ %CLIENTVERSIONC2R% (
echo ^> %MACHINE% has Office products installed with v%CLIENTVERSIONC2R%
echo ^> You cannot install out of band releases.
echo ^> Change the install source or client version.
echo. & pause & echo. & GOTO STARTINSTALLC2R
)

call :DEPLOYSETUPC2R
SET "REGADDINSTALLC2R=Y"
SET "UPDATEPATHC2R=%DEPLOYDIR%\%SETUPSOURCE%\%BITNESS%"
call :XMLINSTALLC2R
call :SETAUTHORC2R
call :SCHEDCREATEC2R
echo.
echo You are about to install %PRODUCT_ID% on %MACHINE%^^!
echo.
choice /c YN /N /M "Do you want to continue? <Y/N>"
IF ERRORLEVEL 2 call :CLEANUPC2R & GOTO STARTODT2016
IF ERRORLEVEL 1 GOTO RUNINSTALLC2R

:RUNINSTALLC2R
call :REGADDPATHC2R
call :SCHEDRUNC2R
call :SETUPWAITC2R
echo. & pause & echo. & GOTO STARTINSTALLC2R

::~~~~~~~~~~~~::
::Call scripts::
::~~~~~~~~~~~~::
:SETMACHINEC2R
SET /p "MACHINE=Enter computer name: "
ping -n 1 -4 %MACHINE% | FIND "TTL=" >nul
IF %ERRORLEVEL% EQU 1 (
echo. & echo Could not connect to "%MACHINE%" & echo. & pause & GOTO STARTODT2016
) else (
GOTO REMOTEREGC2R
)
:REMOTEREGC2R
REG QUERY "\\%MACHINE%\HKLM" >>nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
echo. & echo %MACHINE% is online, but cannot connect to remote registry^^! & echo. & pause & GOTO STARTODT2016
) else (
echo.
)
GOTO EOF

:SETBITNESSC2R
echo [3] 32 bit
echo [6] 64 bit
echo.
choice /c 36 /N /M ">"
IF ERRORLEVEL 2 SET "BITNESS=x64" & GOTO EOF
IF ERRORLEVEL 1 SET "BITNESS=x86" & GOTO EOF

:SETSOURCEC2R
IF %SETUPTYPE% EQU Update (
FOR /f "tokens=*" %%A in ('dir /a:d /b "%DEPLOYDIR%\ALL\%PLATFORMC2R%\Office\Data" 2^>nul') do (SET "SRCVERALLC2R=%%A")
FOR /f "tokens=*" %%B in ('dir /a:d /b "%DEPLOYDIR%\Beta\%PLATFORMC2R%\Office\Data" 2^>nul') do (SET "SRCVERBETAC2R=%%B")
FOR /f "tokens=*" %%C in ('dir /a:d /b "%DEPLOYDIR%\Alpha\%PLATFORMC2R%\Office\Data" 2^>nul') do (SET "SRCVERALPHAC2R=%%C")
)
IF %SETUPTYPE% EQU Install (
FOR /f "tokens=*" %%A in ('dir /a:d /b "%DEPLOYDIR%\ALL\%BITNESS%\Office\Data" 2^>nul') do (SET "SRCVERALLC2R=%%A")
FOR /f "tokens=*" %%B in ('dir /a:d /b "%DEPLOYDIR%\Beta\%BITNESS%\Office\Data" 2^>nul') do (SET "SRCVERBETAC2R=%%B")
FOR /f "tokens=*" %%C in ('dir /a:d /b "%DEPLOYDIR%\Alpha\%BITNESS%\Office\Data" 2^>nul') do (SET "SRCVERALPHAC2R=%%C")
)
IF NOT DEFINED SRCVERALLC2R SET "SRCVERALLC2R=empty"
IF NOT DEFINED SRCVERBETAC2R SET "SRCVERBETAC2R=empty"
IF NOT DEFINED SRCVERALPHAC2R SET "SRCVERALPHAC2R=empty"
echo [1] ALL - %SRCVERALLC2R%
echo [2] Beta - %SRCVERBETAC2R%
echo [3] Alpha - %SRCVERALPHAC2R%
echo.
choice /c 123 /N /M ">"
IF ERRORLEVEL 3 SET "SETUPSOURCE=Alpha" & GOTO EOF
IF ERRORLEVEL 2 SET "SETUPSOURCE=Beta" & GOTO EOF
IF ERRORLEVEL 1 SET "SETUPSOURCE=All" & GOTO EOF

:REGQUERYC2R
SET "PLATFORMC2R="
SET "O365PRODUCTC2R="
SET "VISIOPRODUCTC2R="
SET "CLIENTVERSIONC2R="
SET "PROJECTPRODUCTC2R="
FOR /f "usebackq tokens=3" %%Z in (`REG QUERY "\\%MACHINE%\%KEYC2RCONFIG%" /v "%VALUEVERC2R%" 2^>nul ^| find "%VALUEVERC2R%"`) do (SET "CLIENTVERSIONC2R=%%Z")
FOR /f "usebackq tokens=3" %%Y in (`REG QUERY "\\%MACHINE%\%KEYC2RPOLICY%" /v "%VALUEUPDATEPATH%" 2^>nul ^| find "%VALUEUPDATEPATH%"`) do (SET "UPDATEPATHC2R=%%Y")
FOR /f "usebackq tokens=3" %%X in (`REG QUERY "\\%MACHINE%\%KEYC2RCONFIG%" /v "%VALUEUPDATEURL%" 2^>nul ^| find "%VALUEUPDATEURL%"`) do (SET "UPDATEURLC2R=%%X")
FOR /f "usebackq tokens=3" %%W in (`REG QUERY "\\%MACHINE%\%KEYC2RCONFIG%" /v "%VALUEPLATFORM%" 2^>nul ^| find "%VALUEPLATFORM%"`) do (SET "PLATFORMC2R=%%W")
FOR /f "usebackq tokens=3" %%V in (`REG QUERY "\\%MACHINE%\%KEYC2ROFFICE16%" /v "%VALUEVERSION%" 2^>nul ^| find "%VALUEVERSION%"`) do (SET "O365VERSIONC2R=%%V")
FOR /f "usebackq tokens=3" %%U in (`REG QUERY "\\%MACHINE%\%KEYC2RPROJECT16%" /v "%VALUEVERSION%" 2^>nul ^| find "%VALUEVERSION%"`) do (SET "PROJECTVERSIONC2R=%%U")
FOR /f "usebackq tokens=3" %%T in (`REG QUERY "\\%MACHINE%\%KEYC2RVISIO16%" /v "%VALUEVERSION%" 2^>nul ^| find "%VALUEVERSION%"`) do (SET "VISIOVERSIONC2R=%%T")
FOR /f "usebackq tokens=3" %%S in (`REG QUERY "\\%MACHINE%\%KEYC2ROFFICE16%" /v "%VALUENAME%" 2^>nul ^| find "%VALUENAME%"`) do (SET "O365PRODUCTC2R=%%S")
FOR /f "usebackq tokens=3" %%R in (`REG QUERY "\\%MACHINE%\%KEYC2RPROJECT16%" /v "%VALUENAME%" 2^>nul ^| find "%VALUENAME%"`) do (SET "PROJECTPRODUCTC2R=%%R")
FOR /f "usebackq tokens=3" %%Q in (`REG QUERY "\\%MACHINE%\%KEYC2RVISIO16%" /v "%VALUENAME%" 2^>nul ^| find "%VALUENAME%"`) do (SET "VISIOPRODUCTC2R=%%Q")
IF NOT DEFINED PROJECTPRODUCTC2R SET "PROJECTPRODUCTC2R=NUL"
IF NOT DEFINED CLIENTVERSIONC2R SET "CLIENTVERSIONC2R=NUL"
IF NOT DEFINED VISIOPRODUCTC2R SET "VISIOPRODUCTC2R=NUL"
IF NOT DEFINED O365PRODUCTC2R SET "O365PRODUCTC2R=NUL"
IF NOT DEFINED PLATFORMC2R SET "PLATFORMC2R=NUL"
GOTO EOF

:VERSIONREPORTC2R
IF %CLIENTVERSIONC2R% EQU NUL GOTO SETUPDATEPATHC2R
echo ^> Version on client
echo %CLIENTVERSIONC2R% %PLATFORMC2R%
echo.
:SETUPDATEPATHC2R
IF NOT DEFINED UPDATEPATHC2R (
SET "UPDATEPATHC2R=%UPDATEURLC2R%" & GOTO VIRGINPATHC2R
) else (
echo ^> Update path
echo %UPDATEPATHC2R%
echo. & GOTO EOF
)
:VIRGINPATHC2R
echo ^> Update path (original - not GPO assigned)
IF DEFINED UPDATEPATHC2R echo %UPDATEPATHC2R%
echo. & GOTO EOF

:DOWNGRADEC2R
SET "SETUPTYPE=Downgrade"
IF %SKIPREGPATH% EQU Y GOTO SKIPREGPATHC2R
echo.
call :REGADDPATHC2R
:SKIPREGPATHC2R
call :OC2RSWITCH1
FOR /f "tokens=*" %%T in ('dir /a:d /b "%UPDATEPATHC2R%\Office\Data" 2^>nul') do (SET "DOWNGRADEVERSIONC2R=%%T")
call :DEPLOYSETUPC2R
call :SETAUTHORC2R
call :SCHEDCREATEC2R
echo.
choice /c YN /N /M "Run downgrade now? <Y/N>"
IF ERRORLEVEL 2 call :CLEANUPC2R & GOTO STARTODT2016
IF ERRORLEVEL 1 GOTO DOWNGRADENOWCR2
:DOWNGRADENOWCR2
call :SCHEDRUNC2R
call :UPDATEWAITC2R
echo. & pause & echo. & GOTO EOF

:REGADDPATHC2R
reg add "\\%MACHINE%\%KEYC2RPOLICY%" /v updatepath /t REG_SZ /d "%UPDATEPATHC2R%" /f >nul 2>nul
IF %ERRORLEVEL% NEQ 0 echo ^> Failed to write registry value 1 of 2 & echo.
reg add "\\%MACHINE%\%KEYC2RCONFIG%" /v UpdateUrl /t REG_SZ /d "%UPDATEPATHC2R%" /f >nul 2>nul
IF %ERRORLEVEL% NEQ 0 echo ^> Failed to write registry value 2 of 2
IF %REGADDINSTALLC2R% NEQ Y (
IF %ERRORLEVEL% EQU 0 echo ^> Path changed successfully
)
GOTO EOF

:OC2RSWITCH1
echo.
choice /c YN /N /M "Prompt user? <Y/N>"
IF ERRORLEVEL 2 SET "SWITCHPROMPTC2R=False" & echo. & GOTO OC2RSWITCH2
IF ERRORLEVEL 1 SET "SWITCHPROMPTC2R=True" & echo. & GOTO OC2RSWITCH2
:OC2RSWITCH2
choice /c YN /N /M "Show progress? <Y/N>"
IF ERRORLEVEL 2 SET "SWITCHDISPLAYC2R=False" & echo. & GOTO OC2RSWITCH3
IF ERRORLEVEL 1 SET "SWITCHDISPLAYC2R=True" & echo. & GOTO OC2RSWITCH3
:OC2RSWITCH3
choice /c YN /N /M "Force close Office apps? <Y/N>"
IF ERRORLEVEL 2 SET "SWITCHFORCEC2R=False" & echo. & GOTO EOF
IF ERRORLEVEL 1 SET "SWITCHFORCEC2R=True" & echo. & GOTO EOF


:DEPLOYSETUPC2R
MD "\\%MACHINE%\%SETUPDIRNET%" 2>nul
(
echo Set WshShell = CreateObject^("WScript.Shell" ^)
echo WshShell.Run chr^(34^) ^& "%SETUPDIRLOC%\Office2016_%SETUPTYPE%.bat" ^& Chr^(34^), 0
echo Set WshShell = Nothing
) > "\\%MACHINE%\%SETUPDIRNET%\Office2016_%SETUPTYPE%.vbs" 2>&1

IF NOT EXIST "\\%MACHINE%\%SETUPDIRNET%\Office2016_%SETUPTYPE%.vbs" echo. & echo Failed to create setup files on %MACHINE%^^! & echo. & pause & call :CLEANUPC2R & GOTO STARTODT2016
IF %SETUPTYPE% EQU Update GOTO DEPLOYUPDATEC2R
IF %SETUPTYPE% EQU Downgrade GOTO DEPLOYDOWNGRADEC2R

(
echo SETLOCAL ENABLEDELAYEDEXPANSION
IF %SETUPTYPE% EQU Install (
echo "%DEPLOYDIR%\setup.exe" /configure "%XMLDIR%\Config%SETUPTYPE%%PRODUCT_ID%%SETUPSOURCE%_%BITNESS%.xml"
)
IF %SETUPTYPE% EQU Uninstall (
echo "%DEPLOYDIR%\setup.exe" /configure "%XMLDIR%\Config%SETUPTYPE%%PRODUCT_ID%.xml"
)
echo REG QUERY "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%PRODUCT_ID% - en-us" /V "DisplayName"
IF %SETUPTYPE% EQU Install (
echo IF %%ERRORLEVEL%% EQU 0 ^(
echo SET "SETUPCHECK=0"
echo ^) else ^(
echo SET "SETUPCHECK=1"
echo ^)
)
IF %SETUPTYPE% EQU Uninstall (
echo IF %%ERRORLEVEL%% EQU 0 ^(
echo SET "SETUPCHECK=1"
echo ^) else ^(
echo SET "SETUPCHECK=0"
echo ^)
)
echo ^(
echo echo Setup completed at %time:~-11,2%:%time:~-8,2%:%time:~-5,2%
echo echo.
echo IF %%SETUPCHECK%% EQU 0 ^(
echo echo Successfully %SETUPTYPE%ed %PRODUCT_ID%
echo ^) else ^(
echo echo FAILED^^! to %SETUPTYPE% %PRODUCT_ID%
echo ^)
echo ^) ^> %LOGDIR%\%MACHINE%_%SETUPTYPE%%PRODUCT_ID%%BITNESS%-%CURDATE%.log
echo TIMEOUT /t 5 /nobreak ^>nul
echo SCHTASKS /delete /tn "Office2016_%SETUPTYPE%" /f
echo CD /d C:\
echo START cmd /c rd /s /q "%SETUPDIRLOC%"
echo :EOF
) > "\\%MACHINE%\%SETUPDIRNET%\Office2016_%SETUPTYPE%.bat" 2>&1
GOTO EOF

:DEPLOYUPDATEC2R
(
echo "%OC2RCLIENT%" /update user updatepromptuser=%SWITCHPROMPTC2R% forceappshutdown=%SWITCHFORCEC2R% displaylevel=%SWITCHDISPLAYC2R%
echo SET "EXITCODE=%%ERRORLEVEL%%"
echo IF %%EXITCODE%% EQU 0 ^(
echo SET "RESULT=SUCCESS"
echo ^) else ^(
echo SET "RESULT=FAIL"
echo ^)
echo ^(
echo echo OfficeC2RClient executed at %time:~-11,2%:%time:~-8,2%:%time:~-5,2%
echo echo Exit code: %%EXITCODE%%
echo echo %%RESULT%%
echo ^) ^> %LOGDIR%\%MACHINE%_%SETUPTYPE%%PLATFORMC2R%-%CURDATE%.log
echo TIMEOUT /t 5 /nobreak ^>nul
echo SCHTASKS /delete /tn "Office2016_%SETUPTYPE%" /f
echo CD /d C:\
echo START cmd /c rd /s /q "%SETUPDIRLOC%"
echo :EOF
) > "\\%MACHINE%\%SETUPDIRNET%\Office2016_%SETUPTYPE%.bat" 2>&1
GOTO EOF

:DEPLOYDOWNGRADEC2R
(
echo "%OC2RCLIENT%" /update user updatetoversion=%DOWNGRADEVERSIONC2R% updatepromptuser=%SWITCHPROMPTC2R% forceappshutdown=%SWITCHFORCEC2R% displaylevel=%SWITCHDISPLAYC2R%
echo SET "EXITCODE=%%ERRORLEVEL%%"
echo IF %%EXITCODE%% EQU 0 ^(
echo SET "RESULT=SUCCESS"
echo ^) else ^(
echo SET "RESULT=FAIL"
echo ^)
echo ^(
echo echo OfficeC2RClient executed at %time:~-11,2%:%time:~-8,2%:%time:~-5,2%
echo echo Exit code: %%EXITCODE%%
echo echo %%RESULT%%
echo ^) ^> %LOGDIR%\%MACHINE%_%SETUPTYPE%%PLATFORMC2R%-%CURDATE%.log
echo TIMEOUT /t 5 /nobreak ^>nul
echo SCHTASKS /delete /tn "Office2016_%SETUPTYPE%" /f
echo CD /d C:\
echo START cmd /c rd /s /q "%SETUPDIRLOC%"
echo :EOF
) > "\\%MACHINE%\%SETUPDIRNET%\Office2016_%SETUPTYPE%.bat" 2>&1
GOTO EOF

:XMLDOWNLOADC2R
IF %BITNESS% EQU x64 SET "BITNESS=64"
IF %BITNESS% EQU x86 SET "BITNESS=86"
(
echo ^<Configuration^>
echo     ^<Add SourcePath="%DEPLOYDIR%\Download\%PRODUCT_ID%-%C2RCHANNEL%Channelx%BITNESS%_%CURDATE%" OfficeClientEdition="%BITNESS%" Channel="%C2RCHANNEL%" OfficeMgmtCOM="TRUE"^>
echo         ^<Product ID="%PRODUCT_ID%"^>
echo             ^<Language ID="en-us" /^>
echo         ^</Product^>
echo     ^</Add^>
echo     ^<Logging Level="Standard" Path="%%WINDIR%%\Temp" /^>
echo ^</Configuration^>
) > "%XMLDIR%\Config%SETUPTYPE%%PRODUCT_ID%-%C2RCHANNEL%_x%BITNESS%.xml"
IF %BITNESS% EQU 64 SET "BITNESS=x64"
IF %BITNESS% EQU 86 SET "BITNESS=x86"
GOTO EOF

:XMLUNINSTALLC2R
(
echo ^<Configuration^>
IF %REMOVEALLC2R% EQU True (
echo     ^<Remove All="%REMOVEALLC2R%" /^>
)
IF %REMOVEALLC2R% EQU False (
echo     ^<Remove^>
echo        ^<Product ID="%PRODUCT_ID%"^>
echo            ^<Language ID="en-us" /^>
echo        ^</Product^>
echo     ^</Remove^>
)
echo     ^<Display Level="None" AcceptEULA="TRUE" /^>
echo     ^<Logging Level="Standard" Path="%%WINDIR%%\Temp" /^>
echo     ^<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" /^>
echo ^</Configuration^>
) > "%XMLDIR%\Config%SETUPTYPE%%PRODUCT_ID%.xml"
GOTO EOF

:XMLINSTALLC2R
IF %BITNESS% EQU x64 SET "BITNESS=64"
IF %BITNESS% EQU x86 SET "BITNESS=86"
(
echo ^<Configuration^>
echo     ^<Add SourcePath="%DEPLOYDIR%\%SETUPSOURCE%\x%BITNESS%" OfficeClientEdition="%BITNESS%" OfficeMgmtCOM="TRUE"^>
echo         ^<Product ID="%PRODUCT_ID%"^>
echo             ^<Language ID="en-us" /^>
echo         ^</Product^>
echo     ^</Add^>
echo     ^<Updates Enabled="TRUE" UpdatePath="%DEPLOYDIR%\%SETUPSOURCE%\x%BITNESS%" /^>
echo     ^<Display Level="None" AcceptEULA="TRUE" /^>
echo     ^<Logging Level="Standard" Path="%%WINDIR%%\Temp" /^>
echo     ^<Property Name="AUTOACTIVATE" Value="0" /^>
echo     ^<Property Name="FORCEAPPSHUTDOWN" Value="FALSE" /^>
echo     ^<Property Name="SharedComputerLicensing" Value="0" /^>
echo     ^<Property Name="PinIconsToTaskbar" Value="FALSE" /^>
echo ^</Configuration^>
) > "%XMLDIR%\Config%SETUPTYPE%%PRODUCT_ID%%SETUPSOURCE%_x%BITNESS%.xml"
IF %BITNESS% EQU 64 SET "BITNESS=x64"
IF %BITNESS% EQU 86 SET "BITNESS=x86"
GOTO EOF

:SETAUTHORC2R
SET "psCommand=powershell -Command "$pword = read-host 'Enter password for %USERDOMAIN%\%USERNAME%' -AsSecureString ; ^
    $BSTR=[System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pword); ^
        [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)""
FOR /f "usebackq delims=" %%p in (`%psCommand%`) do set PWD=%%p
title Office 2016 Deployment Tool
GOTO EOF

:SCHEDCREATEC2R
SCHTASKS /Create /S "%MACHINE%" /RU "%USERDOMAIN%\%USERNAME%" /RP "%PWD%" /SC ONEVENT /EC Application /MO *[System/EventID=969] /TN "Office2016_%SETUPTYPE%" /TR "%SETUPDIRLOC%\Office2016_%SETUPTYPE%.vbs" /F /RL HIGHEST >>"%LOGDIR%\Temp\OC2R-SCHCreate%SETUPTYPE%.log" 2>&1
FINDSTR /i /c:"password is incorrect" /c:"bad password" "%LOGDIR%\Temp\OC2R-SCHCreate%SETUPTYPE%.log" >nul
IF %ERRORLEVEL% EQU 0 echo. & echo Authentication FAILED^^! & echo. & call :CLEANUPC2R & pause & GOTO STARTODT2016
FINDSTR /i /c:"access is denied" "%LOGDIR%\Temp\OC2R-SCHCreate%SETUPTYPE%.log" >nul
IF %ERRORLEVEL% EQU 0 echo. & echo Authentication FAILED^^! & echo. & call :CLEANUPC2R & pause & GOTO STARTODT2016
FINDSTR /c:"SUCCESS" "%LOGDIR%\Temp\OC2R-SCHCreate%SETUPTYPE%.log" >nul
IF %ERRORLEVEL% EQU 0 (
SET NUL=NUL
) else (
TYPE "%LOGDIR%\Temp\OC2R-SCHCreate%SETUPTYPE%.log" & echo. & echo Error on task create. & echo. & pause & GOTO STARTODT2016
)
GOTO EOF

:SCHEDRUNC2R
SCHTASKS /Run /S "%MACHINE%" /U "%USERDOMAIN%\%USERNAME%" /P "%PWD%" /TN "Office2016_%SETUPTYPE%" >>"%LOGDIR%\Temp\OC2R-SCHRun%SETUPTYPE%.log" 2>&1
FINDSTR /c:"SUCCESS" "%LOGDIR%\Temp\OC2R-SCHRun%SETUPTYPE%.log" >nul
IF %ERRORLEVEL% EQU 0 (
echo. & echo Task started at %time:~-11,2%:%time:~-8,2%:%time:~-5,2% please wait...
) else (
TYPE "%LOGDIR%\Temp\OC2R-SCHRun%SETUPTYPE%.log" & echo. & echo Error on task run. & echo. & pause & GOTO STARTODT2016
)
GOTO EOF

:UPDATEWAITC2R
IF EXIST "%LOGDIR%\%MACHINE%_%SETUPTYPE%%PLATFORMC2R%-%CURDATE%.log" (
del /q /f "%LOGDIR%\%MACHINE%_%SETUPTYPE%%PLATFORMC2R%-%CURDATE%.log"
)
:UPDATEWAITLOOP1
echo|SET /p=.
timeout /t 2 /nobreak >nul
SET /a U = U + 1
IF NOT EXIST "%LOGDIR%\%MACHINE%_%SETUPTYPE%%PLATFORMC2R%-%CURDATE%.log" (
GOTO UPDATEWAITLOOP1
) else (
GOTO UPDATEWAITLOOP2
)
:UPDATEWAITLOOP2
echo|SET /p=.
timeout /t 2 /nobreak >nul
SET /a U = U + 1
tasklist /s "%MACHINE%" /svc /fi "IMAGENAME eq OfficeClickToRun.exe" | FINDSTR /i /c:"OfficeClickToRun.exe" | FINDSTR /i /c:"N/A" >nul
IF %ERRORLEVEL% EQU 0 (
GOTO UPDATEWAITLOOP2
) else (
echo. & echo. & echo ^> Process completed at %time:~-11,2%:%time:~-8,2%:%time:~-5,2%
)
GOTO EOF

:SETUPWAITC2R
IF EXIST "%LOGDIR%\%MACHINE%_%SETUPTYPE%%PRODUCT_ID%%BITNESS%-%CURDATE%.log" (
del /q /f "%LOGDIR%\%MACHINE%_%SETUPTYPE%%PRODUCT_ID%%BITNESS%-%CURDATE%.log"
)
:SETUPWAITLOOP
echo|SET /p=.
timeout /t 2 /nobreak >nul
SET /a S = S + 1
IF NOT EXIST "%LOGDIR%\%MACHINE%_%SETUPTYPE%%PRODUCT_ID%%BITNESS%-%CURDATE%.log" (
GOTO SETUPWAITLOOP
) else (
echo. & echo. & TYPE "%LOGDIR%\%MACHINE%_%SETUPTYPE%%PRODUCT_ID%%BITNESS%-%CURDATE%.log"
)
GOTO EOF

:CLEANUPC2R
IF %MACHINE% EQU %COMPUTERNAME% (SCHTASKS /delete /tn "Office2016_%SETUPTYPE%" /f >>nul 2>&1)
IF %MACHINE% EQU localhost (SCHTASKS /delete /tn "Office2016_%SETUPTYPE%" /f >>nul 2>&1)
IF %MACHINE% EQU 127.0.0.1 (SCHTASKS /delete /tn "Office2016_%SETUPTYPE%" /f >>nul 2>&1)
SCHTASKS /S "%MACHINE%" /U "%USERDOMAIN%\%USERNAME%" /P "%PWD%" /delete /tn "Office2016_%SETUPTYPE%" /f >>nul 2>&1
RD /s /q \\%MACHINE%\%SETUPDIRNET% >nul 2>nul

:EOF
