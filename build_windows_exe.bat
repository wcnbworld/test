@echo off
setlocal

REM Build single-file exe for Windows
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

pyinstaller --noconfirm --clean --onefile --name PatentClassifier launcher.py

echo.
echo Build completed. EXE location:
echo dist\PatentClassifier.exe
pause
