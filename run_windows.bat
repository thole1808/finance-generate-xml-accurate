@echo off
setlocal

python --version >nul 2>&1
if errorlevel 1 (
  echo Python belum terinstall atau belum masuk PATH.
  echo Install Python 3.11/3.12 lalu centang "Add python.exe to PATH".
  pause
  exit /b 1
)

python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Gagal install dependency.
  pause
  exit /b 1
)

python app.py
endlocal
