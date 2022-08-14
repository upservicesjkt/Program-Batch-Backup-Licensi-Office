@ECHO off
cls
cls
mode con:cols=50 lines=25
COLOR 08
set _time=%time:~0,8%
echo Time : [ %_time% ] @ [ %date% ]
ECHO.

title Backup Licensi

set nama=
set /p nama=Masukan Nama :
set merek=
set /p merek=Masukan Merek :
set SerialLicensi=
set /p SerialLicensi=Masukan Licensi :
:Office
ECHO.
ECHO 1. Office Retail
ECHO 2. Office VL
ECHO.
set menu=
set /p menu=Pilih Type Office.
if '%menu%'=='1' goto Retail
if '%menu%'=='2' goto VL
ECHO "%menu%" is not valid, try again
ECHO.
goto Office

:Retail
cls
ECHO.
ECHO 1. Office 2010
ECHO 2. Office 2013 
ECHO 3. Office 2016
ECHO 4. Office 2019
ECHO.
set menu=
set /p menu=Pilih Type Office.
if '%menu%'=='1' goto 2010R
if '%menu%'=='2' goto 2013R
if '%menu%'=='3' goto 2016R
if '%menu%'=='4' goto 2019R
ECHO "%menu%" is not valid, try again
ECHO.
goto Retail

:2010R
set versi=Office 2010 Retail
goto backupoffice
:2013R
set versi=Office 2013 Retail
goto backupoffice
:2016R
set versi=Office 2016 Retail
goto backupoffice
:2019R
set versi=Office 2019 Retail
goto backupoffice

:VL
cls
ECHO.
ECHO 1. Office 2010
ECHO 2. Office 2013 
ECHO 3. Office 2016
ECHO 4. Office 2019
ECHO.
set menu=
set /p menu=Pilih Type Office.
if '%menu%'=='1' goto 2010VL
if '%menu%'=='2' goto 2013VL
if '%menu%'=='3' goto 2016VL
if '%menu%'=='4' goto 2019VL
ECHO "%menu%" is not valid, try again
ECHO.
goto VL

:2010VL
set versi=Office 2010 VL
goto backupoffice
:2013VL
set versi=Office 2013 VL
goto backupoffice
:2016VL
set versi=Office 2016 VL
goto backupoffice
:2019VL
set versi=Office 2019 VL
goto backupoffice
:backupoffice
mkdir "C:\Backup"
echo %SerialLicensi% >> "C:\Backup\Licensi.txt"
echo Extrak spp to C:\Windows\System32 >> "C:\Backup\Read me ....txt"
xcopy /s /i "C:\Windows\System32\spp" "C:\Backup\spp"
copy "C:\Result.txt" "C:\Backup\"

ECHO.:Backup
Call :GetSerialNumber

for /f "tokens=3,2,4 delims=/- " %%x in ("%date%") do set d=%%y%%x%%z
set data=%d%
Echo zipping...
"%~dp0\7-Zip\7z.exe" a -tzip "C:\%versi% Pro Plus (%merek% SN %SerialNumber% %nama%).zip" "C:\Backup\Licensi.txt" "C:\Backup\Read me ....txt" "C:\Backup\spp" "C:\Result.txt"
rmdir /s /q "C:\Backup"
echo Done!
::***********************************************
:GetSerialNumber
for /f "tokens=2 delims==" %%a in ('
    wmic csproduct get identifyingnumber /value
') do for /f "delims=" %%b in ("%%a") do (
    Set "SerialNumber=%%b" 
)
exit /b
::***********************************************