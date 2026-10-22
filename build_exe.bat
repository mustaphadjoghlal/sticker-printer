@echo off
chcp 65001 >nul
echo Installing requirements...
python -m pip install -r requirements.txt pyinstaller
echo Building StickerPrinter.exe ...
python -m PyInstaller --noconsole --onefile --name StickerPrinter app.py
echo.
echo Done. The program is in the dist folder: dist\StickerPrinter.exe
pause
