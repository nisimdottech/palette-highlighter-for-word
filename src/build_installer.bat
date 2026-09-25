@echo off
setlocal
cd /d "%~dp0.."

powershell -NoProfile -ExecutionPolicy Bypass -File src\embed_vba.ps1
if errorlevel 1 exit /b 1

uv run python src\assemble_template.py
if errorlevel 1 exit /b 1

uvx --from pyinstaller pyinstaller --noconfirm --clean --onefile --windowed --name PaletteHighlighterForWord --add-data "%CD%\build\PaletteHighlighterForWord.dotm;." --distpath "%CD%\dist" --workpath "%CD%\build\pyinstaller" --specpath "%CD%\build\pyinstaller" "%CD%\src\installer.py"
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

echo Build completed successfully.
endlocal
