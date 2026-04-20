@echo off
chcp 65001 > nul
echo [1/2] Installing PyInstaller...
pip install pyinstaller

echo [2/2] Building EbookCapture.exe...
python -m PyInstaller --noconfirm --onefile --windowed --name "EbookCapture" --icon "icon.ico" --collect-all dxcam --collect-all comtypes --hidden-import cv2 --hidden-import numpy --hidden-import keyboard --hidden-import img2pdf capture.py

echo.
if exist dist\EbookCapture.exe (
    echo Done! dist\EbookCapture.exe
) else (
    echo Build failed.
)
pause
