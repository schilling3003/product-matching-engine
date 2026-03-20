@echo off

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ProductMatcher.spec del ProductMatcher.spec

echo Building the executable with PyInstaller (simple version)...
pyinstaller --name "ProductMatcher" ^
    --onefile ^
    --add-data "app.py;." ^
    --add-data "src;src" ^
    --add-data ".streamlit;.streamlit" ^
    --hidden-import=streamlit ^
    --hidden-import=streamlit.web.cli ^
    --collect-submodules=streamlit ^
    simple_launcher.py

echo Build process finished.
pause
