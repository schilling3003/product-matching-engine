@echo off

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ProductMatcher.spec del ProductMatcher.spec

echo Building the executable with PyInstaller (DEBUG MODE - Console)...
pyinstaller --name "ProductMatcher" ^
    --onefile ^
    --console ^
    --add-data "app.py;." ^
    --hidden-import=streamlit ^
    --hidden-import=streamlit.web.cli ^
    --hidden-import=streamlit.runtime ^
    --hidden-import=streamlit.runtime.scriptrunner ^
    --hidden-import=streamlit.runtime.state ^
    --hidden-import=streamlit.components.v1 ^
    --hidden-import=thefuzz ^
    --hidden-import=thefuzz.fuzz ^
    --hidden-import=thefuzz.process ^
    --hidden-import=sklearn ^
    --hidden-import=sklearn.feature_extraction.text ^
    --hidden-import=sklearn.metrics.pairwise ^
    --hidden-import=numpy ^
    --hidden-import=multiprocessing ^
    --hidden-import=concurrent.futures ^
    --collect-all=streamlit ^
    --collect-all=altair ^
    --collect-all=plotly ^
    launcher.py

echo Build process finished.
echo Testing the executable...
.\dist\ProductMatcher.exe
pause
