@echo off
setlocal

echo ==============================================
echo   JPG TO EXCEL CONVERTER - Build Script
echo ==============================================

set VENV_DIR=venv
set TESSERACT_SRC="C:\Program Files\Tesseract-OCR"
set APP_NAME=JPG_To_Excel_Converter
set DIST_DIR=dist\%APP_NAME%

if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo Virtual environment not found. Creating one...
    python -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo Failed to create virtual environment. Is Python installed?
        exit /b 1
    )
)

call "%VENV_DIR%\Scripts\activate.bat"

echo.
echo Installing/updating dependencies...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo Failed to install dependencies.
    exit /b 1
)

echo.
echo Cleaning previous build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "%APP_NAME%.spec" del /q "%APP_NAME%.spec"

echo.
echo Building EXE with PyInstaller...
pyinstaller --noconfirm --onedir --windowed --name "%APP_NAME%" main.py
if errorlevel 1 (
    echo PyInstaller build failed.
    exit /b 1
)

if not exist "%DIST_DIR%" (
    echo Build failed - dist folder not found.
    exit /b 1
)

echo.
echo Bundling Tesseract-OCR runtime (so the target PC needs nothing installed)...
if exist %TESSERACT_SRC% (
    xcopy /e /i /y %TESSERACT_SRC% "%DIST_DIR%\Tesseract-OCR" >nul
    echo Tesseract-OCR bundled successfully.
) else (
    echo WARNING: Tesseract-OCR not found at %TESSERACT_SRC%.
    echo Copy a Tesseract-OCR installation into "%DIST_DIR%\Tesseract-OCR"
    echo manually, otherwise the built EXE will not be able to run OCR.
)

echo.
echo ==============================================
echo   Build complete
echo   EXE:    %DIST_DIR%\%APP_NAME%.exe
echo   Ship the whole "%DIST_DIR%" folder as-is.
echo ==============================================

endlocal
