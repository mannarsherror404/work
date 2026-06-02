@echo off
setlocal

cd /d "%~dp0"

echo Creating virtual environment...

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 -m venv .venv
) else (
    python -m venv .venv
)

if errorlevel 1 (
    echo Failed to create virtual environment.
    echo Please install Python 3 and make sure it is added to PATH.
    pause
    exit /b 1
)

echo Activating virtual environment...
call ".venv\Scripts\activate.bat"

echo Upgrading pip...
python -m pip install --upgrade pip

echo Installing project dependencies...
python -m pip install -r requirements.txt

echo Installing PyInstaller...
python -m pip install pyinstaller

echo Building Windows EXE...
python -m PyInstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name DataMaskTransformer ^
    --add-data "assets;assets" ^
    app.py

if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Build complete.
echo Your app is here:
echo %CD%\dist\DataMaskTransformer.exe
echo.
pause
