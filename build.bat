@echo off
echo Installing PyInstaller...
call venv\Scripts\activate
pip install pyinstaller

echo.
echo Building executable...
pyinstaller --onefile --name XLtoJSON --clean __main__.py

echo.
echo Build complete! Executable is in dist\XLtoJSON.exe
echo.
pause
