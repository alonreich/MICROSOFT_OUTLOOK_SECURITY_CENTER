@echo off
taskkill /F /IM electron.exe /T >nul 2>&1
start "" "node_modules\electron\dist\electron.exe" . %*
exit
