@echo off

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ProductMatcher.spec del ProductMatcher.spec

echo Building the executable with PyInstaller (fixed version)...
pyinstaller --name "ProductMatcher" ^
    --onefile ^
    --noconsole ^
    --add-data "app.py;." ^
    --add-data "src;src" ^
    --add-data ".streamlit;.streamlit" ^
    --exclude-module=tkinter ^
    --exclude-module=matplotlib ^
    --exclude-module=IPython ^
    --exclude-module=jupyter ^
    --exclude-module=notebook ^
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
    --hidden-import=openpyxl ^
    --collect-submodules=streamlit ^
    --collect-data=streamlit ^
    --copy-metadata=streamlit ^
    --copy-metadata=pandas ^
    --copy-metadata=numpy ^
    --copy-metadata=scikit-learn ^
    launcher.py

echo Build process finished.
pause
