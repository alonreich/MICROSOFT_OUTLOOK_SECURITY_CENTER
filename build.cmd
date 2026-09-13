@echo off
setlocal
echo [DeskGuard Build] Starting packaging process...
npm run package
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Build failed with exit code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
echo.
echo [SUCCESS] Packaging complete: Output in dist\
endlocal
