@ECHO OFF
:: ---------------------------------------------------------------
:: Version 2.0
::
:: NAME    : KMS Activation
::
:: AUTHOR  : Kasper Stad
:: GITHUB  : http://github.kasperstad.dk
::
:: DATE    : 2016-01-10
::
:: COMMENT : Force activate Microsoft Windows and Office using KMS
:: ---------------------------------------------------------------

:: KMS Server informations, please correct to the correct values
SET kmsip=kms.contoso.com
SET kmsport=1688

:: Change the design of the console window
set title=Key Management Service Activation 
COLOR 17
title %title%

:: Check if the script is running as administrator, else stop script
net session >nul 2>&1
if %errorLevel% == 0 (
    GOTO execute
) else (
    COLOR 4E
        echo You need to run this file as Administrator!
        echo Right click the file, choose Run as administrator.
    pause
    exit.
)

:: Set KMS Server
:execute
pause
    echo Setting KMS server to: %kmsip%:%kmsport%
    %windir%\System32\slmgr.vbs /skms %kmsip%:%kmsport%
    echo Done..
    goto activate
pause

:: Activate Windows and Office
:activate

echo Detecting if office is installed, and activating it...

:: Office 2010
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" set SCRIPT="C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" && goto officeactivate
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" set SCRIPT="C:\Program Files\Microsoft Office\Office14\ospp.vbs" && goto officeactivate

:: Office 2013
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" set SCRIPT="C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" && goto officeactivate
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" set SCRIPT="C:\Program Files\Microsoft Office\Office15\ospp.vbs" && goto officeactivate

:: Office 2016
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" set SCRIPT="C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" && goto officeactivate
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" set SCRIPT="C:\Program Files\Microsoft Office\Office16\ospp.vbs" && goto officeactivate

:: Attempt to activate Microsoft Office
:officeactivate

IF DEFINED SCRIPT (

    echo Office is installed... activating...
    %windir%\System32\cscript.exe %SCRIPT% /act
    echo Office activated...

) ELSE (

    echo No office version detected... 

)

goto windowsactivate


:: Attempt to activate Microsoft Windows
:windowsactivate

echo Activating Windows...
%windir%\System32\slmgr.vbs /ato
echo Windows activation completed

pause