@echo off
title ShortsMaker - one time setup
cd /d "%~dp0"
where python >nul 2>nul || (echo Python nahi mila. https://www.python.org/downloads/ se Python 3.10+ install karo ^(installer mein "Add python.exe to PATH" tick karna^). & pause & exit /b 1)
echo Virtual environment bana raha hoon...
python -m venv venv || (echo venv nahi bana. & pause & exit /b 1)
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt || (echo Install fail hua, internet check karo. & pause & exit /b 1)
echo.
echo Whisper model download kar raha hoon - sirf pehli baar, lagbhag 500 MB...
python -c "from faster_whisper import WhisperModel; WhisperModel('small', device='cpu', compute_type='int8')"
echo.
echo Setup poora hua. Ab apni video ko run.bat par drag-and-drop karo.
pause
