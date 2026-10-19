@echo off
title ShortsMaker
cd /d "%~dp0"
if not exist venv\Scripts\python.exe (echo Pehle setup.bat chalao. & pause & exit /b 1)
set "VIDEO=%~1"
if "%VIDEO%"=="" set /p VIDEO=Video file ka path daalo (ya file ko yahan drag karo): 
set VIDEO=%VIDEO:"=%
set /p CLIPS=Kitni clips chahiye? [Enter = 5]: 
if "%CLIPS%"=="" set CLIPS=5
set /p VLANG=Video ki bhasha? hi / en / Enter = auto: 
set LANGARG=
if not "%VLANG%"=="" set LANGARG=--lang %VLANG%
venv\Scripts\python.exe shorts_maker.py "%VIDEO%" --clips %CLIPS% %LANGARG% %2 %3 %4 %5 %6 %7 %8 %9
echo.
pause
