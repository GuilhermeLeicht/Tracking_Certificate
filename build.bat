@echo off
REM Gera dist\Tracking Certificate\  (execute na pasta do projeto, com o venv ativo)
pip install pyinstaller
pyinstaller --noconfirm --clean --windowed ^
  --name "Tracking Certificate" ^
  --icon "icon.ico" ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pyodbc ^
  --hidden-import plyer.platforms.win.notification ^
  --exclude-module PySide6.QtWebEngineCore ^
  --exclude-module PySide6.QtQml ^
  --exclude-module PySide6.Qt3DCore ^
  main.py
pause
