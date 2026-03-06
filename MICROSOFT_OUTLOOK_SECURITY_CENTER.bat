@echo off
set "ELECTRON_RUN_AS_NODE="
start "" "node_modules\.bin\electron.cmd" . --no-sandbox
exit