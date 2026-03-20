@echo off

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ProductMatcher.spec del ProductMatcher.spec

echo Building DIAGNOSTIC executable with enhanced launcher...
pyinstaller --name "ProductMatcher_Diagnostic" ^
    --onefile ^
    --console ^
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
    launcher_diagnostic.py

echo Build process finished.
echo.
echo DIAGNOSTIC VERSION CREATED: dist\ProductMatcher_Diagnostic.exe
echo.
echo Instructions for your work computer:
echo 1. Copy ProductMatcher_Diagnostic.exe to your work computer
echo 2. Open Command Prompt as Administrator
echo 3. Navigate to the folder containing the executable
echo 4. Run: ProductMatcher_Diagnostic.exe
echo 5. The console will show detailed diagnostic information
echo 6. Take a screenshot or copy the error messages
echo.
pause