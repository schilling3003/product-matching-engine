@echo off
echo Running ProductMatcher.exe...
echo.
cd /d "C:\Users\User\CodeOutTool\product-matching-engine\dist"
ProductMatcher.exe
echo.
echo Exit code: %ERRORLEVEL%
echo.
pause
