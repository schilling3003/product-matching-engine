@echo off

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ProductMatcher.spec del ProductMatcher.spec

echo Building DIRECTORY-BASED executable (more compatible)...
pyinstaller --name "ProductMatcher" ^
    --onedir ^
    --windowed ^
    --add-data "app.py;." ^
    --add-data ".streamlit;.streamlit" ^
    --hidden-import=streamlit ^
    --hidden-import=streamlit.web.cli ^
    --hidden-import=streamlit.runtime ^
    --hidden-import=streamlit.runtime.scriptrunner ^
    --hidden-import=streamlit.runtime.state ^
    --hidden-import=streamlit.components.v1 ^
    --hidden-import=thefuzz ^
    --hidden-import=thefuzz.fuzz ^
    --hidden-import=thefuzz.process ^
    --hidden-import=numpy ^
    --hidden-import=multiprocessing ^
    --hidden-import=concurrent.futures ^
    --hidden-import=sklearn ^
    --hidden-import=sklearn.feature_extraction.text ^
    --hidden-import=sklearn.metrics.pairwise ^
    --collect-all=streamlit ^
    --collect-all=altair ^
    --collect-all=plotly ^
    launcher.py

echo Build process finished.
echo.
echo DIRECTORY VERSION CREATED: dist\ProductMatcher\
echo.
echo To deploy:
echo 1. Copy the entire dist\ProductMatcher folder to your work computer
echo 2. Run ProductMatcher.exe from inside that folder
echo.
echo This version is often more compatible with corporate environments.
echo.
pause
